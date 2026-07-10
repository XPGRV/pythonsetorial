"""
coletar_usda_ers_ldp_forecasts.py

Automatiza a coleta dos forecasts do USDA ERS (série: Livestock, Dairy, and Poultry Outlook).

O foco para "beef forecast" é o download **U.S. red meat and poultry forecasts (XLSX)**,
que contém a planilha "RMPFORE" com:
  - Production (million pounds) — linha "Beef"
  - Per capita disappearance (retail pounds) — linha "Beef"
  - Market prices (ex.: Steers 5-area Direct, etc.)

Este script:
  1) Descobre os releases (um por mês) na página da série
  2) Para cada release, extrai os 3 links de download (PDF / Dairy / Red meat)
  3) Baixa o XLSX "U.S. red meat and poultry forecasts" (prioritário)
  4) Converte a tabela "RMPFORE" para formato tidy (linhas = métrica; colunas = período)
  5) Exporta CSV(s) para uma pasta local

Dependências (instale via pip):
  - requests
  - beautifulsoup4
  - pandas
  - openpyxl

Uso (exemplos):
  python coletar_usda_ers_ldp_forecasts.py --latest
  python coletar_usda_ers_ldp_forecasts.py --months 6
  python coletar_usda_ers_ldp_forecasts.py --all
"""

from __future__ import annotations

import datetime as dt
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup

import openpyxl
from openpyxl.styles import Alignment, Font


SERIES_URL = "https://www.ers.usda.gov/publications?series=LDPM"

# Endpoint JSON que a própria página da série consome via JavaScript.
# (descoberto em drupalSettings.ers_blocks.requestUrl = "/api/publications/v1.0?")
# Retorna sempre os releases mais recentes da série, sem precisar de pubids fixos.
API_URL = "https://www.ers.usda.gov/api/publications/v1.0"
SERIES_CODE = "LDPM"

# URLs secundárias de fallback.
# Nota: a maioria das páginas ERS é JS-rendered e não retorna dados via requests.
# Mantemos apenas a página da série antiga como tentativa.
_EXTRA_SOURCES: list[str] = [
    "https://www.ers.usda.gov/publications/pub-details/?pubid=37653",
]

# Pubids conhecidos para releases que a página da série não renderiza no HTML.
# Chave: (ano, mês); Valor: pubid do ERS.
# Adicione aqui sempre que precisar fixar um release específico.
_KNOWN_PUBIDS: dict[tuple[int, int], int] = {
    (2025, 12): 113593,  # December 2025
    (2026,  1): 113688,  # January 2026
    (2026,  2): 113843,  # February 2026
    (2026,  3): 113953,  # March 2026
    (2026,  4): 114064,  # April 2026
}

XLSX_NAME = "BeefUS.xlsm"

# Nome exibido do download (varia pouco ao longo do tempo)
DOWNLOAD_LABEL_RED_MEAT = "U.S. red meat and poultry forecasts"
DOWNLOAD_LABEL_RED_MEAT_ALT = "US red meat and poultry forecasts"


@dataclass(frozen=True)
class Release:
    title: str
    date: dt.date | None
    url: str


@dataclass(frozen=True)
class ReleaseDownloads:
    release: Release
    pdf_url: str | None
    dairy_xlsx_url: str | None
    red_meat_xlsx_url: str | None


def _session() -> requests.Session:
    s = requests.Session()
    s.headers.update(
        {
            "User-Agent": "Mozilla/5.0 (compatible; BeefBR/1.0; +https://ers.usda.gov/)",
            "Accept-Language": "en-US,en;q=0.9",
        }
    )
    return s


def _parse_release_date(text: str) -> dt.date | None:
    # Ex.: "4/15/2026" (M/D/YYYY)
    text = (text or "").strip()
    m = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{4})", text)
    if not m:
        return None
    mm, dd, yyyy = map(int, m.groups())
    try:
        return dt.date(yyyy, mm, dd)
    except ValueError:
        return None


def _discover_release_by_pubid_range(
    s: requests.Session,
    year: int,
    month: int,
    known_pubids: dict[tuple[int, int], int],
) -> Release | None:
    """
    Quando o HTML da série não traz um release, estima o pubid pelo passo médio
    dos pubids conhecidos e testa um range ao redor do estimado.
    Retorna o Release se encontrado, None caso contrário.
    """
    sorted_known = sorted(known_pubids.items())  # [((y,m), pubid), ...]
    if not sorted_known:
        return None

    # Calcula passo médio entre pubids consecutivos (em pubids/mês)
    steps = []
    for i in range(1, len(sorted_known)):
        (y1, m1), p1 = sorted_known[i - 1]
        (y2, m2), p2 = sorted_known[i]
        months_diff = (y2 - y1) * 12 + (m2 - m1)
        if months_diff > 0:
            steps.append((p2 - p1) / months_diff)
    avg_step = sum(steps) / len(steps) if steps else 110.0

    # Referência: pubid mais recente conhecido
    (ref_y, ref_m), ref_pubid = sorted_known[-1]
    months_from_ref = (year - ref_y) * 12 + (month - ref_m)
    estimate = int(ref_pubid + months_from_ref * avg_step)

    # Busca saindo do estimado em espiral: est, est+3, est-3, est+6, est-6...
    _EN_MON = {
        "january": 1, "february": 2, "march": 3, "april": 4,
        "may": 5, "june": 6, "july": 7, "august": 8,
        "september": 9, "october": 10, "november": 11, "december": 12,
    }
    base = "https://www.ers.usda.gov/publications/pub-details?pubid="
    candidates = [estimate]
    for delta in range(3, 180, 3):
        candidates += [estimate + delta, estimate - delta]

    for pubid in candidates:
        if pubid <= 0:
            continue
        try:
            html = s.get(f"{base}{pubid}", timeout=20).text
            if "pub-details" not in html:
                continue
            soup = BeautifulSoup(html, "html.parser")
            h1 = soup.find("h1")
            if not h1:
                continue
            title = h1.get_text(" ", strip=True)
            if not title.startswith("Livestock, Dairy, and Poultry Outlook"):
                continue
            tm = re.search(r":\s*([A-Za-z]+)\s+(\d{4})", title)
            if not tm:
                continue
            t_mon = _EN_MON.get(tm.group(1).lower())
            t_yr = int(tm.group(2))
            if t_mon == month and t_yr == year:
                known_pubids[(year, month)] = pubid  # cacheia para próximas chamadas
                return Release(title=title, date=dt.date(year, month, 15), url=f"{base}{pubid}")
        except Exception:
            continue
    return None


def _scrape_releases_from_soup(soup: BeautifulSoup, base_url: str) -> list[Release]:
    """
    Extrai releases de um BeautifulSoup já parseado.
    Tenta primeiro a seção div#releases (cards); depois varre todos os links pub-details.
    """
    found: list[Release] = []

    # Caminho 1: seção de cards "Releases"
    releases_root = soup.select_one("div#releases")
    if releases_root:
        for li in releases_root.select("ul.usa-card-group > li"):
            a = li.select_one("a[href*='pub-details']")
            if not a:
                continue
            title = a.get_text(" ", strip=True)
            if not title.startswith("Livestock, Dairy, and Poultry Outlook"):
                continue
            url = urljoin(base_url, a.get("href") or "")
            date_val: dt.date | None = None
            p = li.find("p")
            if p:
                date_val = _parse_release_date(p.get_text(" ", strip=True))
            found.append(Release(title=title, date=date_val, url=url))

    # Caminho 2: varre TODOS os links pub-details na página (inclui tabelas, listas, etc.)
    for a in soup.select("a[href*='pub-details']"):
        title = a.get_text(" ", strip=True)
        href = a.get("href") or ""
        if not title.startswith("Livestock, Dairy, and Poultry Outlook"):
            continue
        url = urljoin(base_url, href)
        date_val = None
        parent_text = a.parent.get_text(" ", strip=True) if a.parent else ""
        m = re.search(r"(\d{1,2}/\d{1,2}/\d{4})", parent_text)
        if m:
            date_val = _parse_release_date(m.group(1))
        # tenta inferir pelo título se não achou data
        if date_val is None:
            tm = re.search(r":\s*([A-Za-z]+)\s+(\d{4})", title)
            if tm:
                en_to_month = {
                    "january": 1, "february": 2, "march": 3, "april": 4,
                    "may": 5, "june": 6, "july": 7, "august": 8,
                    "september": 9, "october": 10, "november": 11, "december": 12,
                }
                mon = en_to_month.get(tm.group(1).lower())
                yr = int(tm.group(2))
                if mon:
                    try:
                        date_val = dt.date(yr, mon, 15)  # aproximado: dia 15
                    except ValueError:
                        pass
        found.append(Release(title=title, date=date_val, url=url))

    return found


def _list_releases_from_api(series_code: str = SERIES_CODE, size: int = 24) -> list[Release]:
    """
    Fonte primária: consome o mesmo endpoint JSON que a página da série usa.
    Não depende de JS rendering nem de pubids fixos — sempre traz os releases
    mais recentes da série (o `id` do JSON é o próprio pubid).
    """
    s = _session()
    r = s.get(
        API_URL,
        params={"series": series_code, "size": size},
        headers={"Accept": "application/json"},
        timeout=60,
    )
    r.raise_for_status()
    data = r.json()

    out: list[Release] = []
    for row in data.get("rows", []):
        title = (row.get("title") or "").strip()
        if not title.startswith("Livestock, Dairy, and Poultry Outlook"):
            continue

        date_val: dt.date | None = None
        rd = row.get("releaseDate")
        if rd:
            try:
                date_val = dt.date.fromisoformat(str(rd)[:10])
            except ValueError:
                date_val = None

        rel_url = (row.get("url") or "").strip()
        if rel_url.startswith("/"):
            rel_url = "https://www.ers.usda.gov" + rel_url
        if not rel_url:
            # fallback: monta a partir do id (= pubid)
            pubid = row.get("id")
            if pubid:
                rel_url = f"https://www.ers.usda.gov/publications/{pubid}"
        if not rel_url:
            continue

        out.append(Release(title=title, date=date_val, url=rel_url))

    return out


def list_releases(series_url: str = SERIES_URL) -> list[Release]:
    """
    Lista os releases da série LDP Outlook.

    Fonte primária: API JSON (`/api/publications/v1.0?series=LDPM`), que é o
    mesmo endpoint consumido pela página da série. Independe de pubids fixos.

    Fallback (caso a API mude/caia): scraping do HTML + pubids conhecidos.
    """
    # --- Fonte primária: API JSON ---
    try:
        api_rel = _list_releases_from_api()
        if api_rel:
            api_rel.sort(
                key=lambda r: (r.date is None, r.date or dt.date.min), reverse=True
            )
            return api_rel
    except Exception as e:
        print(f"[WARN] API ERS indisponível ({e}); usando fallback de scraping.", flush=True)

    # --- Fallback: scraping de múltiplas fontes ---
    s = _session()
    all_candidates: list[Release] = []

    # Fonte 1: página da série
    try:
        html = s.get(series_url, timeout=60).text
        soup = BeautifulSoup(html, "html.parser")
        all_candidates.extend(_scrape_releases_from_soup(soup, series_url))

        # Tenta seguir link "View all releases" ou similar na mesma página
        for a in soup.find_all("a", href=True):
            txt = a.get_text(" ", strip=True).lower()
            if any(kw in txt for kw in ("view all", "all releases", "see all", "mais releases")):
                view_all_url = urljoin(series_url, a["href"])
                try:
                    va_html = s.get(view_all_url, timeout=60).text
                    va_soup = BeautifulSoup(va_html, "html.parser")
                    all_candidates.extend(_scrape_releases_from_soup(va_soup, view_all_url))
                except Exception:
                    pass
    except Exception:
        pass

    # Fontes extras: busca/listagem ERS
    for extra_url in _EXTRA_SOURCES:
        try:
            html = s.get(extra_url, timeout=60).text
            soup = BeautifulSoup(html, "html.parser")
            all_candidates.extend(_scrape_releases_from_soup(soup, extra_url))
        except Exception:
            pass

    # Pubids conhecidos: busca os que ainda não estiverem na lista
    found_month_years = {
        (r.date.year, r.date.month) for r in all_candidates if r.date
    }
    base = "https://www.ers.usda.gov/publications/pub-details?pubid="
    mutable_known = dict(_KNOWN_PUBIDS)  # cópia mutável para caching de discoveries
    for (yr, mon), pubid in mutable_known.items():
        if (yr, mon) in found_month_years:
            continue
        url = f"{base}{pubid}"
        try:
            html = s.get(url, timeout=15).text
            soup = BeautifulSoup(html, "html.parser")
            h1 = soup.find("h1")
            title = h1.get_text(" ", strip=True) if h1 else ""
            if not title.startswith("Livestock, Dairy, and Poultry Outlook"):
                continue
            date_val: dt.date | None = dt.date(yr, mon, 15)
            all_candidates.append(Release(title=title, date=date_val, url=url))
            found_month_years.add((yr, mon))
            print(f"  [pubid={pubid}] {title}", flush=True)
        except Exception:
            pass

    # Deduplica por URL, mantendo a primeira ocorrência (que costuma ter mais info)
    uniq: dict[str, Release] = {}
    for r in all_candidates:
        uniq.setdefault(r.url, r)

    out = list(uniq.values())
    out.sort(key=lambda r: (r.date is None, r.date or dt.date.min), reverse=True)
    return out


def get_release_downloads(release: Release) -> ReleaseDownloads:
    s = _session()
    html = s.get(release.url, timeout=60).text
    soup = BeautifulSoup(html, "html.parser")

    def _find_download_under_heading(heading_text: str) -> str | None:
        # Estrutura típica:
        # <li ...>
        #   <h3 ...>HEADING</h3>
        #   <a href="...xlsx|pdf">Download ...</a>
        # </li>
        h = soup.find(lambda tag: tag.name in {"h2", "h3", "h4"} and tag.get_text(" ", strip=True) == heading_text)
        if not h:
            return None
        container = h.find_parent(["li", "div", "section"]) or h.parent
        if not container:
            return None
        for a in container.select("a[href]"):
            href = (a.get("href") or "").strip()
            if not href:
                continue
            full = urljoin(release.url, href)
            if re.search(r"\.(pdf|xlsx|xls)(?:$|\?)", full, flags=re.I):
                return full
        return None

    # Preferimos identificar os 3 downloads pelo heading (robusto, pois o texto do link costuma ser "Download XLSX/PDF").
    pdf_url = _find_download_under_heading(release.title)  # o PDF costuma ter o mesmo título do release
    dairy_url = _find_download_under_heading("Dairy Forecasts") or _find_download_under_heading("Dairy forecasts")
    red_meat_url = _find_download_under_heading("U.S. red meat and poultry forecasts") or _find_download_under_heading(
        "US red meat and poultry forecasts"
    )

    return ReleaseDownloads(release=release, pdf_url=pdf_url, dairy_xlsx_url=dairy_url, red_meat_xlsx_url=red_meat_url)


def download_file(url: str, dest_path: Path) -> Path:
    dest_path.parent.mkdir(parents=True, exist_ok=True)
    s = _session()
    with s.get(url, stream=True, timeout=120) as r:
        r.raise_for_status()
        with open(dest_path, "wb") as f:
            for chunk in r.iter_content(chunk_size=1024 * 256):
                if chunk:
                    f.write(chunk)
    return dest_path


def parse_rmpfore_to_tidy(xlsx_path: Path, release: Release) -> pd.DataFrame:
    """
    Converte a planilha "RMPFORE" (layout ERS) para formato tidy:
      release_title, release_date, release_url, section, item, year, period, value
    onde period é I/II/III/IV/Annual.
    """
    df = pd.read_excel(xlsx_path, sheet_name="RMPFORE", header=None)

    # Padrão observado:
    # row 1 (index 1): ... 2024 ... 2025 ... 2026 ...
    # row 2 (index 2): ... I II III IV Annual ...
    year_row = df.iloc[1].tolist()
    period_row = df.iloc[2].tolist()

    years: list[int | None] = []
    cur_year: int | None = None
    for v in year_row:
        y = None
        if isinstance(v, (int, float)) and pd.notna(v):
            iv = int(v)
            if 1900 <= iv <= 2100:
                y = iv
        if y is not None:
            cur_year = y
        years.append(cur_year)

    periods: list[str | None] = []
    for v in period_row:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            periods.append(None)
            continue
        s = str(v).strip()
        if s in {"I", "II", "III", "IV", "Annual"}:
            periods.append(s)
        else:
            periods.append(None)

    # Colunas de dados: onde temos year e period
    data_cols: list[int] = [
        i for i, (y, p) in enumerate(zip(years, periods)) if (y is not None and p is not None)
    ]

    def cell_to_str(x) -> str | None:
        if x is None:
            return None
        if isinstance(x, float) and pd.isna(x):
            return None
        s = str(x).strip()
        return s or None

    current_section: str | None = None
    rows = []

    # Dados começam tipicamente em index 4
    for ridx in range(3, len(df)):
        first = cell_to_str(df.iat[ridx, 0] if df.shape[1] > 0 else None)
        if first is None:
            continue

        # "Section title" costuma ser uma linha com texto e demais colunas vazias
        # Ex.: "Production, million pounds"
        others = df.iloc[ridx, 1:].dropna()
        if len(others) == 0:
            current_section = first
            continue

        # Linha de item (ex.: Beef, Pork, ...)
        item = first
        if current_section is None:
            # fallback: evita perder dados se a planilha mudar um pouco
            current_section = "Unknown"

        for c in data_cols:
            y = years[c]
            p = periods[c]
            val = df.iat[ridx, c]
            if val is None or (isinstance(val, float) and pd.isna(val)):
                continue
            rows.append(
                {
                    "release_title": release.title,
                    "release_date": release.date.isoformat() if release.date else None,
                    "release_url": release.url,
                    "section": current_section,
                    "item": item,
                    "year": int(y) if y is not None else None,
                    "period": p,
                    "value": float(val) if isinstance(val, (int, float)) else val,
                }
            )

    out = pd.DataFrame(rows)
    return out


def extract_production_quarters_with_forecast_flag(
    xlsx_path: Path, release: Release, item: str = "Beef"
) -> pd.DataFrame:
    """
    Extrai "Production, million pounds" / item especificado do sheet RMPFORE,
    preservando o sinal de forecast via formatação (negrito).

    Parâmetros:
      item : nome exato da linha na planilha (ex.: "Beef", "Broilers")

    Retorna tidy com colunas:
      release_title, release_date, release_url, item, ref, year, period, value, is_forecast

    Onde:
      - period ∈ {I, II, III, IV, Annual}
      - ref = "YYYYQn" (ou "YYYY" para Annual)
      - is_forecast = True se célula estiver em negrito (font.bold == True)
    """
    wb = openpyxl.load_workbook(xlsx_path, read_only=False, data_only=True)
    try:
        ws = wb["RMPFORE"]

        # Header: linha 2 = anos; linha 3 = I/II/III/IV/Annual (1-based)
        year_row = 2
        period_row = 3

        years: dict[int, int] = {}
        cur_year: int | None = None
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=year_row, column=c).value
            y: int | None = None
            if isinstance(v, (int, float)) and v is not None:
                iv = int(v)
                if 1900 <= iv <= 2100:
                    y = iv
            if y is not None:
                cur_year = y
            if cur_year is not None:
                years[c] = cur_year

        periods: dict[int, str] = {}
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=period_row, column=c).value
            if v is None:
                continue
            s = str(v).strip()
            if s in {"I", "II", "III", "IV", "Annual"}:
                periods[c] = s

        data_cols = [c for c in range(1, ws.max_column + 1) if (c in years and c in periods)]

        # Localiza seção "Production, million pounds" e a linha "Beef" dentro dessa seção.
        section_target = "Production, million pounds"
        current_section: str | None = None
        beef_row: int | None = None
        for r in range(1, ws.max_row + 1):
            first = ws.cell(row=r, column=1).value
            if first is None or (isinstance(first, str) and not first.strip()):
                continue
            first_s = str(first).strip()

            # linha de seção: as demais colunas geralmente vazias
            any_other = False
            for c in range(2, min(ws.max_column, 10) + 1):
                if ws.cell(row=r, column=c).value not in (None, ""):
                    any_other = True
                    break
            if not any_other:
                current_section = first_s
                continue

            if current_section == section_target and first_s.lower() == item.lower():
                beef_row = r
                break

        if beef_row is None:
            raise RuntimeError(f"Não encontrei a linha '{item}' dentro da seção '{section_target}'.")

        rows = []
        for c in data_cols:
            y = years[c]
            p = periods[c]
            cell = ws.cell(row=beef_row, column=c)
            val = cell.value
            if val is None:
                continue
            bold = bool(cell.font.bold) if cell.font is not None else False

            if p == "Annual":
                ref = f"{y}"
            else:
                qmap = {"I": 1, "II": 2, "III": 3, "IV": 4}
                ref = f"{y}Q{qmap[p]}"

            rows.append(
                {
                    "release_title": release.title,
                    "release_date": release.date.isoformat() if release.date else None,
                    "release_url": release.url,
                    "item": item,
                    "ref": ref,
                    "year": y,
                    "period": p,
                    "value": float(val) if isinstance(val, (int, float)) else val,
                    "is_forecast": bold,
                }
            )

        # --- Derivar trimestres faltantes usando o total anual ---
        # Quando o ERS reporta o anual mas omite algum(ns) trimestre(s),
        # distribuímos o resíduo igualmente pelos trimestres faltantes.
        # Ex.: anual=2026 presente, mas 4Q26 ausente:
        #   4Q26 = (Annual_2026 - 1Q26 - 2Q26 - 3Q26) / 1
        _PERIOD_ORDER = ["I", "II", "III", "IV"]
        _QMAP = {"I": 1, "II": 2, "III": 3, "IV": 4}

        year_periods: dict[int, dict[str, dict]] = {}
        for row in rows:
            year_periods.setdefault(row["year"], {})[row["period"]] = row

        for year, pdata in year_periods.items():
            if "Annual" not in pdata:
                continue
            annual_val = float(pdata["Annual"]["value"])

            present_qs = {p: float(pdata[p]["value"]) for p in _PERIOD_ORDER if p in pdata}
            missing_qs = [p for p in _PERIOD_ORDER if p not in pdata]

            if not missing_qs or not present_qs:
                continue  # nada faltando, ou nenhum trimestre base para derivar

            per_q = (annual_val - sum(present_qs.values())) / len(missing_qs)

            for p in missing_qs:
                ref = f"{year}Q{_QMAP[p]}"
                rows.append(
                    {
                        "release_title": release.title,
                        "release_date": release.date.isoformat() if release.date else None,
                        "release_url": release.url,
                        "item": item,
                        "ref": ref,
                        "year": year,
                        "period": p,
                        "value": per_q,
                        "is_forecast": True,  # valor derivado do anual → forecast
                    }
                )

        return pd.DataFrame(rows)
    finally:
        wb.close()


def build_release_matrix(beef_prod_rows: list[pd.DataFrame]) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Constrói duas matrizes:
      - values: index=ref, columns=release (abreviado), values=value
      - flags:  index=ref, columns=release (abreviado), values=is_forecast (True/False)
    """
    if not beef_prod_rows:
        return pd.DataFrame(), pd.DataFrame()

    all_df = pd.concat(beef_prod_rows, ignore_index=True)
    all_df["release_col"] = (
        all_df["release_title"].str.replace("Livestock, Dairy, and Poultry Outlook:", "", regex=False).str.strip()
    )

    # Ordem das colunas: por data do release (crescente). Se não houver data, cai para o nome.
    release_dates = (
        all_df[["release_col", "release_date"]]
        .drop_duplicates()
        .assign(release_date=lambda d: pd.to_datetime(d["release_date"], errors="coerce"))
    )
    ordered_cols = (
        release_dates.sort_values(["release_date", "release_col"], ascending=[True, True])["release_col"].tolist()
    )

    values = all_df.pivot_table(index="ref", columns="release_col", values="value", aggfunc="first").sort_index()
    flags = all_df.pivot_table(index="ref", columns="release_col", values="is_forecast", aggfunc="first").sort_index()

    # Reaplica a ordem desejada (mantendo apenas colunas existentes)
    ordered_cols = [c for c in ordered_cols if c in values.columns]
    values = values.reindex(columns=ordered_cols)
    flags = flags.reindex(columns=ordered_cols)
    return values, flags


PT_MONTH_ABBR = {
    1: "jan",
    2: "fev",
    3: "mar",
    4: "abr",
    5: "mai",
    6: "jun",
    7: "jul",
    8: "ago",
    9: "set",
    10: "out",
    11: "nov",
    12: "dez",
}


def _release_col_token(release_date_iso: str | None, release_title: str) -> str:
    """
    Converte o release para o token usado no header do Excel do usuário (ex.: 'abr-26').
    Se a data não existir, tenta inferir pelo título; se falhar, retorna o título limpo.
    """
    if release_date_iso:
        try:
            d = dt.date.fromisoformat(release_date_iso)
            return f"{PT_MONTH_ABBR[d.month]}-{str(d.year)[2:]}"
        except Exception:
            pass
    m = re.search(r":\s*([A-Za-z]+)\s+(\d{4})", release_title)
    if m:
        mon = m.group(1).strip().lower()
        year = m.group(2)
        # fallback simples: pega 3 primeiras letras e tenta mapear pt/en
        mon3 = mon[:3]
        en_to_pt = {
            "jan": "jan",
            "feb": "fev",
            "mar": "mar",
            "apr": "abr",
            "may": "mai",
            "jun": "jun",
            "jul": "jul",
            "aug": "ago",
            "sep": "set",
            "oct": "out",
            "nov": "nov",
            "dec": "dez",
        }
        return f"{en_to_pt.get(mon3, mon3)}-{year[2:]}"
    return release_title


def _ref_to_quarter_str(ref: str) -> str | None:
    """
    '2026Q1' -> '1Q26'
    Ignora anuais ('2026').
    """
    m = re.fullmatch(r"(\d{4})Q([1-4])", str(ref))
    if not m:
        return None
    yyyy = m.group(1)
    q = m.group(2)
    return f"{q}Q{yyyy[2:]}"


def _find_cell(ws, needle: str, max_rows: int = 50, max_cols: int = 30) -> tuple[int, int] | None:
    needle_l = needle.strip().lower()
    for r in range(1, min(ws.max_row, max_rows) + 1):
        for c in range(1, min(ws.max_column, max_cols) + 1):
            v = ws.cell(row=r, column=c).value
            if v is None:
                continue
            if str(v).strip().lower() == needle_l:
                return r, c
    return None


def _detect_sheet_colors(
    ws,
    token_to_col: dict[str, int],
    quarter_to_row: dict[str, int],
) -> tuple[str, str]:
    """Detect green and orange RGB (6-char) already used in the sheet cells."""
    green = "00B050"
    orange = "ED7D31"
    for col in token_to_col.values():
        for row in quarter_to_row.values():
            cell = ws.cell(row=row, column=col)
            if cell.value in (None, "") or not cell.font or not cell.font.color:
                continue
            fc = cell.font.color
            if fc.type != "rgb":
                continue
            rgb8 = fc.rgb  # ARGB 8-char from openpyxl
            if rgb8 in ("FF000000", "00000000", "FFFFFFFF", "FF000000"):
                continue
            rgb6 = rgb8[-6:]
            r, g, b = int(rgb6[0:2], 16), int(rgb6[2:4], 16), int(rgb6[4:6], 16)
            if g > r and g > b:
                green = rgb6
            elif r > g and r > b:
                orange = rgb6
    return green, orange


def _neighbor_cell_props(ws, row: int, col: int) -> tuple[Font, str]:
    """
    Retorna (Font, number_format) da célula não-vazia mais próxima à esquerda,
    com fallback para a mesma coluna acima.
    Usar o Font do vizinho como base garante que nome, tamanho e demais
    atributos reflitam o que já está formatado na planilha.
    """
    for c in range(col - 1, 0, -1):
        cell = ws.cell(row=row, column=c)
        if cell.value not in (None, ""):
            return cell.font or Font(), cell.number_format or "General"
    for r in range(row - 1, 0, -1):
        cell = ws.cell(row=r, column=col)
        if cell.value not in (None, ""):
            return cell.font or Font(), cell.number_format or "General"
    return Font(), "General"


def _next_month_token(token: str) -> str | None:
    _months = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
    m = re.fullmatch(r"(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)-(\d{2})", token)
    if not m:
        return None
    idx = _months.index(m.group(1))
    yy = int(m.group(2))
    if idx == 11:
        return f"jan-{(yy + 1) % 100:02d}"
    return f"{_months[idx + 1]}-{yy:02d}"


def _token_to_approx_date(token: str) -> dt.date | None:
    _mon = {"jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
            "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12}
    m = re.fullmatch(r"(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)-(\d{2})", token)
    if not m:
        return None
    try:
        return dt.date(2000 + int(m.group(2)), _mon[m.group(1)], 1)
    except (ValueError, KeyError):
        return None


def update_beefus_production_sheet(
    *,
    beefus_path: Path,
    sheet_name: str,
    values: pd.DataFrame,
    flags: pd.DataFrame,
    release_meta: pd.DataFrame,
    scale_million_to_billion: bool = False,
    remove_sheets_with_prefixes: tuple[str, ...] = ("ERS_",),
) -> None:
    """
    Escreve a matriz (valores e forecast flag) na planilha do usuário:
      - linhas: trimestres na coluna "Quarter" (ex.: 1Q26) — normalmente coluna B
      - colunas: releases por mês (header contém 'abr-26', etc.)
      - estilo: histórico = verde; forecast = laranja

    Observação: a tabela do usuário usa números com 3 casas (ex.: 6.559)
    que equivalem a 'milhões de pounds / 1000' (i.e., bilhões). Então por padrão:
      value_out = value_million / 1000
    """
    if values.empty:
        return

    keep_vba = str(beefus_path).lower().endswith(".xlsm")
    wb = openpyxl.load_workbook(beefus_path, keep_vba=keep_vba)
    try:
        # Remove abas auxiliares antigas (se existirem) para ficar "só Production"
        for pref in remove_sheets_with_prefixes:
            for name in list(wb.sheetnames):
                if name.startswith(pref):
                    try:
                        wb.remove(wb[name])
                    except Exception:
                        pass

        if sheet_name not in wb.sheetnames:
            raise KeyError(f"Aba não encontrada: {sheet_name}")
        ws = wb[sheet_name]

        # Descobre onde está a coluna "Quarter"
        q_cell = _find_cell(ws, "Quarter")
        if not q_cell:
            raise RuntimeError("Não encontrei o header 'Quarter' na aba Production.")
        header_row = q_cell[0]
        quarter_col = q_cell[1]

        # Descobre a linha do header de releases (linha acima do 'Quarter' na maioria dos casos)
        release_header_row = max(1, header_row - 1)

        # Mapeia token do release ('abr-26') -> coluna
        token_to_col: dict[str, int] = {}
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=release_header_row, column=c).value
            if v is None:
                continue
            s = str(v).strip().lower()
            m = re.search(r"\b(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)-\d{2}\b", s)
            if m:
                tok = m.group(0)
                token_to_col[tok] = c

        if not token_to_col:
            raise RuntimeError("Não consegui mapear as colunas de release (ex.: 'abr-26') na linha de header.")

        # Mapeia Quarter string -> linha.
        # Na sua planilha, a coluna "Quarter" é fórmula; então não dá para confiar no data_only.
        # Em vez disso, derivamos o trimestre a partir da coluna "Ano" (datas em A).
        quarter_to_row: dict[str, int] = {}
        ano_col = max(1, quarter_col - 1)  # normalmente A
        for r in range(header_row + 1, ws.max_row + 1):
            v = ws.cell(row=r, column=ano_col).value
            if v is None:
                continue
            if isinstance(v, dt.datetime):
                d = v.date()
            elif isinstance(v, dt.date):
                d = v
            else:
                # tenta parsear texto ISO do Excel
                try:
                    d = dt.datetime.fromisoformat(str(v)).date()
                except Exception:
                    continue
            q = ((d.month - 1) // 3) + 1
            qstr = f"{q}Q{str(d.year)[2:]}"
            quarter_to_row[qstr] = r

        if not quarter_to_row:
            raise RuntimeError("Não encontrei valores de trimestre (ex.: '1Q26') abaixo do header 'Quarter'.")

        # Cores detectadas das células já escritas na planilha
        GREEN, ORANGE = _detect_sheet_colors(ws, token_to_col, quarter_to_row)
        BLACK = "000000"

        # release_meta: release_col (ex: 'April 2026'), release_date (datetime)
        rel_meta = release_meta.copy()
        rel_meta["token"] = rel_meta.apply(
            lambda r: _release_col_token(
                r["release_date"].date().isoformat() if pd.notna(r["release_date"]) else None,
                r["release_title"],
            ),
            axis=1,
        )
        # release_col é a coluna na matriz (ex.: 'April 2026')
        rel_by_col = {r["release_col"]: r["token"] for _, r in rel_meta.iterrows()}

        # Mapa token -> data do release (para filtrar releases futuros)
        today = dt.date.today()
        token_to_reldate: dict[str, dt.date | None] = {
            rm["token"]: (rm["release_date"].date() if pd.notna(rm["release_date"]) else None)
            for _, rm in rel_meta.iterrows()
        }

        def _tok_is_past(tok: str) -> bool:
            d = token_to_reldate.get(tok)
            if d is None:
                # Nenhum release carregado para este token — não preencher via
                # copy-from-left antes de o relatório ser publicado.
                return False
            return d <= today

        # Ordem dos releases (esquerda -> direita) conforme aparecem na planilha
        ordered_tokens_in_sheet = [tok for tok, _c in sorted(token_to_col.items(), key=lambda kv: kv[1])]

        # Escreve célula a célula (somente quando houver dado no ERS)
        for ref in values.index:
            qstr = _ref_to_quarter_str(str(ref))
            if not qstr:
                continue  # ignora Annual
            row = quarter_to_row.get(qstr)
            if not row:
                continue

            for rel_col in values.columns:
                token = rel_by_col.get(rel_col)
                if not token:
                    continue
                col = token_to_col.get(token)
                if not col:
                    continue

                v = values.at[ref, rel_col]
                if pd.isna(v):
                    continue

                out_val = float(v)
                if scale_million_to_billion:
                    out_val = out_val / 1000.0

                # flags pode ser numpy.bool_ (não é o mesmo objeto que True)
                try:
                    is_fc = bool(flags.at[ref, rel_col]) is True
                except Exception:
                    is_fc = False
                cell = ws.cell(row=row, column=col)
                # Na sua aba Production os números parecem estar em "milhões de pounds" como inteiros (ex.: 7106).
                cell.value = int(round(out_val))

                # Herda toda a formatação do vizinho, sobrescrevendo apenas a cor
                base, num_fmt = _neighbor_cell_props(ws, row, col)
                rgb = ORANGE if is_fc else GREEN
                cell.number_format = num_fmt
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(
                    name=base.name,
                    sz=base.sz,
                    b=base.b,
                    i=base.i,
                    charset=base.charset,
                    u=base.u,
                    strike=base.strike,
                    color=openpyxl.styles.Color(rgb=rgb),
                    vertAlign=base.vertAlign,
                    outline=base.outline,
                    shadow=base.shadow,
                    condense=base.condense,
                    extend=base.extend,
                    family=base.family,
                    scheme=base.scheme,
                )

        # Regra extra: dados faltantes "no passado"
        # Se alguma célula ficar vazia em um release já publicado, copia da coluna à esquerda via fórmula (=<left>).
        # Releases futuros (não publicados ainda) são ignorados.
        for qstr, row in quarter_to_row.items():
            # percorre releases na ordem esquerda->direita
            for i in range(1, len(ordered_tokens_in_sheet)):
                tok = ordered_tokens_in_sheet[i]
                # Não preenche releases que ainda não foram publicados
                if not _tok_is_past(tok):
                    continue
                col = token_to_col[tok]
                cell = ws.cell(row=row, column=col)
                if cell.value not in (None, ""):
                    continue
                left_tok = ordered_tokens_in_sheet[i - 1]
                left_col = token_to_col[left_tok]
                left_cell = ws.cell(row=row, column=left_col)
                if left_cell.value in (None, ""):
                    continue
                # fórmula apontando para a célula ao lado — fonte preta, herda formatação do vizinho
                cell.value = f"={left_cell.coordinate}"
                base, num_fmt = _neighbor_cell_props(ws, row, col)
                cell.number_format = num_fmt
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(
                    name=base.name,
                    sz=base.sz,
                    b=base.b,
                    i=base.i,
                    charset=base.charset,
                    u=base.u,
                    strike=base.strike,
                    color=openpyxl.styles.Color(rgb=BLACK),
                    vertAlign=base.vertAlign,
                    outline=base.outline,
                    shadow=base.shadow,
                    condense=base.condense,
                    extend=base.extend,
                    family=base.family,
                    scheme=base.scheme,
                )

        # Para cada release com dados novos escritos, cria o heading do mês seguinte
        # na coluna adjacente (se ainda não existir).
        tokens_written: set[str] = set()
        for ref in values.index:
            qstr = _ref_to_quarter_str(str(ref))
            if not qstr:
                continue
            r = quarter_to_row.get(qstr)
            if not r:
                continue
            for rel_col in values.columns:
                tok = rel_by_col.get(rel_col)
                if not tok or tok not in token_to_col:
                    continue
                if not pd.isna(values.at[ref, rel_col]):
                    tokens_written.add(tok)

        for tok in tokens_written:
            col = token_to_col[tok]
            next_tok = _next_month_token(tok)
            if not next_tok or next_tok in token_to_col:
                continue  # já existe na planilha
            next_col = col + 1
            hdr_cell = ws.cell(row=release_header_row, column=next_col)
            if hdr_cell.value in (None, ""):
                hdr_cell.value = next_tok

        wb.save(beefus_path)
    finally:
        wb.close()


def filter_beef_only(tidy: pd.DataFrame) -> pd.DataFrame:
    """
    Filtro simples: mantém apenas linhas cujo item seja "Beef" (exato).
    """
    return tidy[tidy["item"].astype(str).str.strip().str.lower().eq("beef")].copy()


def run(
    *,
    mode: str,
    months: int | None,
    out_dir: Path,
    keep_downloads: bool,
    write_to: Path | None,
    sheet_prefix: str,
) -> None:
    releases = list_releases()
    if not releases:
        raise RuntimeError("Não encontrei releases na página da série (layout pode ter mudado).")

    if mode == "latest":
        targets = releases[:1]
    elif mode == "months":
        assert months is not None and months > 0
        targets = releases[:months]
    elif mode == "all":
        targets = releases
    else:
        raise ValueError(f"Modo inválido: {mode}")

    downloads_dir = out_dir / "_downloads"
    out_dir.mkdir(parents=True, exist_ok=True)
    downloads_dir.mkdir(parents=True, exist_ok=True)

    all_rows: list[pd.DataFrame] = []
    beef_prod_rows: list[pd.DataFrame] = []
    for rel in targets:
        d = get_release_downloads(rel)
        if not d.red_meat_xlsx_url:
            print(f"[WARN] sem XLSX red meat: {rel.title} ({rel.url})")
            continue

        safe = re.sub(r"[^a-zA-Z0-9._-]+", "_", rel.title).strip("_")
        xlsx_path = downloads_dir / f"{safe}.red_meat.xlsx"
        download_file(d.red_meat_xlsx_url, xlsx_path)

        tidy = parse_rmpfore_to_tidy(xlsx_path, release=rel)
        tidy_beef = filter_beef_only(tidy)

        # Extra: beef production (quarters) + forecast flag (negrito)
        beef_prod = extract_production_quarters_with_forecast_flag(xlsx_path, release=rel)
        beef_prod.to_csv(out_dir / f"{safe}.beef_production_quarters.csv", index=False, encoding="utf-8")
        beef_prod_rows.append(beef_prod)

        # CSVs por release (geral + beef)
        tidy.to_csv(out_dir / f"{safe}.rmpfore.tidy.csv", index=False, encoding="utf-8")
        tidy_beef.to_csv(out_dir / f"{safe}.rmpfore.beef_only.csv", index=False, encoding="utf-8")

        all_rows.append(tidy_beef)

        if not keep_downloads:
            try:
                xlsx_path.unlink(missing_ok=True)
            except Exception:
                pass

    if all_rows:
        merged = pd.concat(all_rows, ignore_index=True)
        merged.to_csv(out_dir / "ERS_LDP_RMPFORE_beef_merged.csv", index=False, encoding="utf-8")

    # Matriz: linhas=trimestres/ano, colunas=release
    values, flags = build_release_matrix(beef_prod_rows)
    if not values.empty:
        # Metadata de releases (para ordenar/nomes)
        all_df = pd.concat(beef_prod_rows, ignore_index=True)
        release_meta = (
            all_df[["release_title", "release_date"]]
            .drop_duplicates()
            .assign(release_date=lambda d: pd.to_datetime(d["release_date"], errors="coerce"))
        )
        release_meta["release_col"] = release_meta["release_title"].str.replace(
            "Livestock, Dairy, and Poultry Outlook:", "", regex=False
        ).str.strip()

        # Se o alvo for o BeefUS.xlsm, NÃO cria abas novas (pedido do usuário):
        # atualiza somente a aba Production.
        if write_to is not None:
            target_name = str(write_to.name).lower()
            if write_to.suffix.lower() == ".xlsm" and "beefus" in target_name:
                update_beefus_production_sheet(
                    beefus_path=write_to,
                    sheet_name="Production",
                    values=values,
                    flags=flags,
                    release_meta=release_meta,
                )
                return

        # Caso contrário, gera o Excel matriz (ou escreve em um workbook existente genérico).
        target_path = write_to if write_to is not None else (out_dir / "ERS_LDP_Beef_Production_Matrix.xlsx")

        keep_vba = str(target_path).lower().endswith(".xlsm")
        if write_to is not None and not target_path.exists():
            raise FileNotFoundError(f"--write-to não encontrado: {target_path}")

        if target_path.exists():
            wb = openpyxl.load_workbook(target_path, keep_vba=keep_vba)
        else:
            wb = openpyxl.Workbook()
            # remove sheet default para evitar lixo
            try:
                default = wb.active
                wb.remove(default)
            except Exception:
                pass

        try:
            values_sheet = f"{sheet_prefix}Values"
            flags_sheet = f"{sheet_prefix}IsForecast"

            def _recreate_sheet(name: str):
                if name in wb.sheetnames:
                    ws_old = wb[name]
                    wb.remove(ws_old)
                return wb.create_sheet(title=name)

            ws_v = _recreate_sheet(values_sheet)
            ws_f = _recreate_sheet(flags_sheet)

            # Escreve matrizes (inclui índice/colunas)
            # Cabeçalho
            ws_v.cell(row=1, column=1, value="ref")
            ws_f.cell(row=1, column=1, value="ref")
            for j, col in enumerate(values.columns, start=2):
                ws_v.cell(row=1, column=j, value=str(col))
                ws_f.cell(row=1, column=j, value=str(col))

            # Linhas
            for i, ref in enumerate(values.index, start=2):
                ws_v.cell(row=i, column=1, value=str(ref))
                ws_f.cell(row=i, column=1, value=str(ref))

            for i, ref in enumerate(values.index, start=2):
                for j, col in enumerate(values.columns, start=2):
                    v = values.at[ref, col]
                    f = flags.at[ref, col]
                    if pd.notna(v):
                        ws_v.cell(row=i, column=j, value=float(v))
                    if f is True:
                        ws_f.cell(row=i, column=j, value=True)
                    elif f is False:
                        ws_f.cell(row=i, column=j, value=False)

            # Negrito em Values quando IsForecast=True
            nrows, ncols = values.shape
            for r in range(nrows):
                for c in range(ncols):
                    flag_cell = ws_f.cell(row=2 + r, column=2 + c).value
                    if flag_cell is True:
                        cell = ws_v.cell(row=2 + r, column=2 + c)
                        cell.font = Font(bold=True)

            wb.save(target_path)
        finally:
            wb.close()


def find_empty_tokens(beefus_path: Path, sheet_name: str = "Production") -> list[str]:
    """
    Abre o BeefUS.xlsm e retorna os tokens (ex.: 'abr-26') cujas colunas
    existem no header mas não têm nenhum dado nos trimestres.
    """
    wb = openpyxl.load_workbook(beefus_path, data_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            raise KeyError(f"Aba não encontrada: {sheet_name}")
        ws = wb[sheet_name]

        q_cell = _find_cell(ws, "Quarter")
        if not q_cell:
            return []
        header_row, quarter_col = q_cell
        release_header_row = max(1, header_row - 1)

        token_to_col: dict[str, int] = {}
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=release_header_row, column=c).value
            if v is None:
                continue
            s = str(v).strip().lower()
            m = re.search(r"\b(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)-\d{2}\b", s)
            if m:
                token_to_col[m.group(0)] = c

        ano_col = max(1, quarter_col - 1)
        quarter_rows: list[int] = [
            r for r in range(header_row + 1, ws.max_row + 1)
            if ws.cell(row=r, column=ano_col).value not in (None, "")
        ]

        if not quarter_rows:
            return []

        empty: list[str] = []
        for tok, col in token_to_col.items():
            has_data = any(
                ws.cell(row=r, column=col).value not in (None, "")
                for r in quarter_rows
            )
            if not has_data:
                empty.append(tok)

        return empty
    finally:
        wb.close()


def main() -> None:
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    import drive_io

    print("Baixando BeefUS.xlsm do Drive...")
    fid = drive_io.find_file_id(XLSX_NAME)
    beefus_path = Path(drive_io.download_to_temp(fid))

    print("Verificando colunas vazias em BeefUS (Production)...")
    missing_beef = find_empty_tokens(beefus_path, sheet_name="Production")
    missing_broiler = []

    missing = sorted(set(missing_beef))
    if not missing:
        print("Nenhuma coluna vazia encontrada. BeefUS ja atualizado.")
        return

    if missing_beef:
        print(f"BeefUS colunas vazias:   {', '.join(missing_beef)}")

    today = dt.date.today()
    print("Buscando lista de releases ERS...")
    all_releases = list_releases()

    # Diagnóstico: mostra todos os releases encontrados
    print(f"  {len(all_releases)} release(s) encontrado(s):")
    for rel in all_releases:
        date_str = rel.date.isoformat() if rel.date else "sem data"
        print(f"    {date_str}  {rel.title}")

    # Mapa inglês -> token PT para fallback por título
    _EN_TO_PT = {
        "january": "jan", "february": "fev", "march": "mar", "april": "abr",
        "may": "mai", "june": "jun", "july": "jul", "august": "ago",
        "september": "set", "october": "out", "november": "nov", "december": "dez",
    }

    def _rel_matches_tok(rel: Release, tok_date: dt.date) -> bool:
        """Bate pelo mês/ano da data; fallback pelo nome do mês no título."""
        if rel.date and rel.date.month == tok_date.month and rel.date.year == tok_date.year:
            return True
        # fallback: "Livestock, Dairy, and Poultry Outlook: December 2025"
        m = re.search(r":\s*([A-Za-z]+)\s+(\d{4})", rel.title)
        if m:
            en_mon = m.group(1).lower()
            yr = int(m.group(2))
            pt_tok = _EN_TO_PT.get(en_mon)
            if pt_tok and yr == tok_date.year:
                mon_num = list(_EN_TO_PT.keys()).index(en_mon) + 1
                if mon_num == tok_date.month:
                    return True
        return False

    targets: list[Release] = []
    for tok in missing:
        tok_date = _token_to_approx_date(tok)
        if not tok_date:
            print(f"[SKIP] {tok}: token inválido.")
            continue
        if tok_date > today:
            print(f"[SKIP] {tok}: release ainda não publicado.")
            continue
        matched = next((r for r in all_releases if _rel_matches_tok(r, tok_date)), None)
        if matched is None:
            # Tenta descobrir o pubid por range (só para este mês específico)
            print(f"[...] {tok}: nao encontrado no HTML, buscando pubid por range...")
            mutable_known = dict(_KNOWN_PUBIDS)
            matched = _discover_release_by_pubid_range(
                _session(), tok_date.year, tok_date.month, mutable_known
            )
        if matched:
            targets.append(matched)
            print(f"[OK]   {tok} -> {matched.title}")
        else:
            print(f"[WARN] {tok}: release nao encontrado na serie ERS.")

    if not targets:
        print("Nenhum release para baixar.")
        return

    cwd = Path(__file__).resolve().parent
    out_dir = cwd / "data_usda_ers"
    downloads_dir = out_dir / "_downloads"
    out_dir.mkdir(parents=True, exist_ok=True)

    beef_prod_rows:    list[pd.DataFrame] = []
    broiler_prod_rows: list[pd.DataFrame] = []

    for rel in targets:
        d = get_release_downloads(rel)
        if not d.red_meat_xlsx_url:
            print(f"[WARN] sem XLSX red meat: {rel.title}")
            continue
        safe = re.sub(r"[^a-zA-Z0-9._-]+", "_", rel.title).strip("_")
        xlsx_path = downloads_dir / f"{safe}.red_meat.xlsx"
        print(f"Baixando XLSX: {rel.title} ...")
        download_file(d.red_meat_xlsx_url, xlsx_path)

        tok = _release_col_token(
            rel.date.isoformat() if rel.date else None, rel.title
        )

        if tok in missing_beef:
            try:
                beef_prod = extract_production_quarters_with_forecast_flag(
                    xlsx_path, release=rel, item="Beef"
                )
                beef_prod_rows.append(beef_prod)
            except Exception as e:
                print(f"[WARN] Beef: {e}")

        if tok in missing_broiler:
            try:
                broiler_prod = extract_production_quarters_with_forecast_flag(
                    xlsx_path, release=rel, item="Broilers"
                )
                broiler_prod_rows.append(broiler_prod)
            except Exception as e:
                print(f"[WARN] Broilers: {e}")

        try:
            xlsx_path.unlink(missing_ok=True)
        except Exception:
            pass

    def _build_and_update(prod_rows: list[pd.DataFrame], target_path: Path, sheet_name: str, label: str) -> None:
        if not prod_rows:
            print(f"Nenhum dado extraído para {label}.")
            return
        values, flags = build_release_matrix(prod_rows)
        all_df = pd.concat(prod_rows, ignore_index=True)
        release_meta = (
            all_df[["release_title", "release_date"]]
            .drop_duplicates()
            .assign(release_date=lambda d: pd.to_datetime(d["release_date"], errors="coerce"))
        )
        release_meta["release_col"] = release_meta["release_title"].str.replace(
            "Livestock, Dairy, and Poultry Outlook:", "", regex=False
        ).str.strip()
        update_beefus_production_sheet(
            beefus_path=target_path,
            sheet_name=sheet_name,
            values=values,
            flags=flags,
            release_meta=release_meta,
        )
        print(f"OK: {label} atualizado -> {target_path}")

    _build_and_update(beef_prod_rows, beefus_path, "Production", "BeefUS")
    if beef_prod_rows:
        drive_io.update_file(fid, str(beefus_path))
        print("OK: BeefUS.xlsm salvo no Drive.")


if __name__ == "__main__":
    main()
