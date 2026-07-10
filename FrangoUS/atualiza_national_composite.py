"""
atualiza_national_composite.py

SOBRE OS DADOS
--------------
Quem reporta: USDA AMS (Agricultural Marketing Service), divisao Livestock,
Poultry and Grain Market News, escritorio de Des Moines (IA).

Frequencia: semanal. O relatorio e publicado toda sexta-feira e cobre as
negociacoes de segunda a sexta daquela semana. Nao ha delay relevante — o
dado ja esta disponivel na API na propria sexta de publicacao.

O que e: o Weekly National Chicken Report (ams_3646) consolida o preco medio
ponderado pelo volume (Wtd Avg) do frango inteiro negociado no mercado spot
dos EUA, na modalidade Domestic Fresh Conventional Delivered. E uma das
referencias de preco spot mais acompanhadas do setor de proteinas nos EUA.

Sao tres series coletadas:
  - National Composite Whole Bird : frango inteiro WOG (Without Giblets,
                                    sem miudezas), todas as categorias de
                                    peso combinadas. E a serie mais ampla
                                    e a mais usada como referencia geral.
  - National Composite WOGS       : subconjunto do Whole Bird, considera
                                    apenas os lotes negociados estritamente
                                    na categoria WOG (exclui RTC).
  - National Composite            : frango inteiro RTC (Ready-to-Cook,
                                    com miudezas), todas as categorias de
                                    peso combinadas.

Unidade: centavos de dolar por libra (USDc/lb).

Historico disponivel na API: a partir de setembro de 2022.

Onde consultar manualmente (PDF semanal e historico):
  https://mymarketnews.ams.usda.gov/viewReport/3646

PDF da semana mais recente (acesso direto):
  https://www.ams.usda.gov/mnreports/ams_3646.pdf

COMO FUNCIONA O CODIGO
----------------------
O codigo se autentica na API MARS (MyMarketNews) do USDA via chave de acesso
e busca o relatorio 3646. A resposta retorna dezenas de linhas com todos os
cortes, regioes e tamanhos negociados na semana. O codigo filtra apenas as
tres linhas de interesse (region="National", size="All Sizes") e grava na
aba NationalComposite do FrangoUS.xlsm a partir da linha 4.

A cada execucao, semanas ja existentes na planilha sao ignoradas
automaticamente, evitando duplicatas.

COMO RODAR
----------
  python atualiza_national_composite.py              # semana atual (padrao)
  python atualiza_national_composite.py --desde 2022 # backfill historico
  python atualiza_national_composite.py --dry-run    # previa sem gravar

Requer: pip install requests openpyxl
"""

from __future__ import annotations

import argparse
import copy
import datetime as dt
import sys
import warnings
from pathlib import Path
from typing import NamedTuple

import openpyxl
import requests

sys.stdout.reconfigure(encoding="utf-8")
warnings.filterwarnings("ignore")


# ── Configuracao ──────────────────────────────────────────────────────────────
import os

# MARS_API_KEY vem do ambiente. Localmente, um .env (fora do git) ao lado
# deste script preenche o ambiente; na nuvem, use GitHub Secrets.
_env_file = Path(__file__).resolve().parent / ".env"
if _env_file.exists():
    for _line in _env_file.read_text(encoding="utf-8-sig").splitlines():
        _line = _line.strip()
        if _line and not _line.startswith("#") and "=" in _line:
            _k, _v = _line.split("=", 1)
            os.environ.setdefault(_k.strip(), _v.strip().strip('"').strip("'"))

MARS_API_KEY = os.environ.get("MARS_API_KEY")
if not MARS_API_KEY:
    raise SystemExit("Defina MARS_API_KEY (variavel de ambiente ou .env).")

DRIVE_ROOT = Path(
    r"G:\.shortcut-targets-by-id\11q1ngdqbySnDMCQSP4fiwPztT6ju2F4Z"
)
XLSX_NAME = "FrangoUS.xlsm"
SHEET_NAME     = "NationalComposite"
FIRST_DATA_ROW = 4   # linhas 1-3 sao cabecalhos

# Indices de coluna (1-based)
COL_QUARTER = 1   # A
COL_MES     = 2   # B
COL_SEMANA  = 3   # C
COL_WB      = 4   # D — National Composite Whole Bird
COL_WOGS    = 5   # E — National Composite WOGS
COL_COMP    = 6   # F — National Composite

MARS_URL = "https://marsapi.ams.usda.gov/services/v1.2/reports/3646"


# ─────────────────────────────────────────────────────────────────────────────
# Estrutura de dados
# ─────────────────────────────────────────────────────────────────────────────

class WeekData(NamedTuple):
    semana:     dt.date
    whole_bird: float
    wogs:       float
    composite:  float


# ─────────────────────────────────────────────────────────────────────────────
# 1. API MARS
# ─────────────────────────────────────────────────────────────────────────────

def localizar_planilha() -> Path:
    """Localiza a planilha sem depender dos nomes das pastas intermediarias."""
    encontrados = list(DRIVE_ROOT.rglob(XLSX_NAME))

    if not encontrados:
        raise FileNotFoundError(
            f"Planilha {XLSX_NAME!r} nao encontrada dentro de {DRIVE_ROOT}"
        )

    if len(encontrados) > 1:
        caminhos = "\n".join(f"  - {path}" for path in encontrados)
        raise RuntimeError(
            f"Foram encontradas varias planilhas chamadas {XLSX_NAME!r}:\n"
            f"{caminhos}"
        )

    return encontrados[0]


def _session() -> requests.Session:
    s = requests.Session()
    s.headers.update({"User-Agent": "Mozilla/5.0 (compatible; FrangoUS-updater/1.0)"})
    return s


def _fetch_mars_range(begin: str, end: str) -> list[dict]:
    """
    Chama a API MARS para um intervalo de datas (begin/end: MM/DD/YYYY).
    Retorna apenas as linhas da secao 'Report Detail'.
    """
    params = {"q": f"report_begin_date={begin}:{end}", "allSections": "true"}
    resp = _session().get(MARS_URL, params=params, auth=(MARS_API_KEY, ""), timeout=60)
    resp.raise_for_status()

    detail_rows: list[dict] = []
    for section in resp.json():
        if section.get("reportSection") == "Report Detail":
            detail_rows.extend(section.get("results", []))
    return detail_rows


def _extract_week_data(rows: list[dict]) -> dict[dt.date, WeekData]:
    """
    Extrai os 3 precos National Composite do detalhe da API.

    Mapeamento (confirmado via inspecao do JSON):
      Whole Bird -> item="ALL Items",         region="National", size="All Sizes"
      WOGS       -> item="WOG",               region="National", size="All Sizes"
      Composite  -> item="RTC Broiler/Fryer", region="National", size="ALL Sizes"
    """
    if not rows:
        return {}

    def _match(row: dict, item: str) -> bool:
        return (
            row.get("item") == item
            and row.get("region") == "National"
            and row.get("size", "").lower() == "all sizes"
        )

    dates: set[str] = {r["report_end_date"] for r in rows if r.get("report_end_date")}

    result: dict[dt.date, WeekData] = {}
    for date_str in sorted(dates):
        day = [r for r in rows if r.get("report_end_date") == date_str]

        wb_row   = next((r for r in day if _match(r, "ALL Items")),         None)
        wogs_row = next((r for r in day if _match(r, "WOG")),               None)
        rtc_row  = next((r for r in day if _match(r, "RTC Broiler/Fryer")), None)

        if wb_row is None or wogs_row is None or rtc_row is None:
            continue

        end_date = dt.datetime.strptime(date_str, "%m/%d/%Y").date()
        result[end_date] = WeekData(
            semana=    end_date,
            whole_bird=float(wb_row["wtd_avg_price"]),
            wogs=      float(wogs_row["wtd_avg_price"]),
            composite= float(rtc_row["wtd_avg_price"]),
        )

    return result


def fetch_recent(weeks_back: int = 2) -> list[WeekData]:
    """Busca as ultimas N semanas (padrao: 2 para garantir que pega a semana atual)."""
    today = dt.date.today()
    begin = today - dt.timedelta(weeks=weeks_back)
    begin_str = begin.strftime("%m/%d/%Y")
    end_str   = today.strftime("%m/%d/%Y")
    print(f"Buscando via API MARS: {begin_str} a {end_str}")
    rows = _fetch_mars_range(begin_str, end_str)
    result = list(_extract_week_data(rows).values())
    print(f"  {len(result)} semana(s) encontrada(s).")
    return result


def fetch_historical(start_year: int) -> list[WeekData]:
    """Busca todos os dados semanais desde start_year em blocos anuais."""
    today   = dt.date.today()
    current = dt.date(start_year, 1, 1)
    result: list[WeekData] = []

    while current <= today:
        end = min(current.replace(year=current.year + 1) - dt.timedelta(days=1), today)
        begin_str = current.strftime("%m/%d/%Y")
        end_str   = end.strftime("%m/%d/%Y")
        print(f"  Buscando {begin_str} a {end_str} ...")
        try:
            rows      = _fetch_mars_range(begin_str, end_str)
            week_data = _extract_week_data(rows)
            result.extend(week_data.values())
            print(f"    {len(week_data)} semana(s) encontrada(s).")
        except Exception as e:
            print(f"  [AVISO] Erro: {e}")
        current = end + dt.timedelta(days=1)

    result.sort(key=lambda w: w.semana)
    return result


# ─────────────────────────────────────────────────────────────────────────────
# 2. Helpers de data
# ─────────────────────────────────────────────────────────────────────────────

def _quarter_str(d: dt.date) -> str:
    q = (d.month - 1) // 3 + 1
    return f"{q}T{str(d.year)[2:]}"


def _mes_date(d: dt.date) -> dt.date:
    return d.replace(day=1)


# ─────────────────────────────────────────────────────────────────────────────
# 3. Atualizacao da planilha
# ─────────────────────────────────────────────────────────────────────────────

def _load_existing_weeks(ws) -> set[dt.date]:
    existing: set[dt.date] = set()
    for row in range(FIRST_DATA_ROW, ws.max_row + 1):
        v = ws.cell(row=row, column=COL_SEMANA).value
        if isinstance(v, dt.datetime):
            existing.add(v.date())
        elif isinstance(v, dt.date):
            existing.add(v)
    return existing


def _next_empty_row(ws) -> int:
    for row in range(FIRST_DATA_ROW, ws.max_row + 2):
        if ws.cell(row=row, column=COL_SEMANA).value is None:
            return row
    return ws.max_row + 1


def _copy_row_format(ws, src_row: int, dst_row: int, n_cols: int) -> None:
    """Copia font, fill, number_format, alignment e border da linha src para dst."""
    for col in range(1, n_cols + 1):
        src = ws.cell(row=src_row, column=col)
        dst = ws.cell(row=dst_row, column=col)
        dst.font          = copy.copy(src.font)
        dst.fill          = copy.copy(src.fill)
        dst.number_format = src.number_format
        dst.alignment     = copy.copy(src.alignment)
        dst.border        = copy.copy(src.border)


def update_sheet(week_list: list[WeekData], dry_run: bool = False) -> dict:
    # Baixa a planilha do Google Drive (em vez do drive montado em G:\).
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    import drive_io
    fid = drive_io.find_file_id(XLSX_NAME)
    path = drive_io.download_to_temp(fid)

    wb = openpyxl.load_workbook(str(path), keep_vba=True, data_only=False)
    try:
        if SHEET_NAME not in wb.sheetnames:
            raise RuntimeError(f"Aba '{SHEET_NAME}' nao encontrada em {path.name}.")
        ws = wb[SHEET_NAME]

        existing = _load_existing_weeks(ws)
        novos    = 0
        skipped  = 0
        log: list[str] = []

        for wd in sorted(week_list, key=lambda w: w.semana):
            if wd.semana in existing:
                skipped += 1
                continue

            row     = _next_empty_row(ws)
            quarter = _quarter_str(wd.semana)
            mes     = _mes_date(wd.semana)

            if not dry_run:
                _copy_row_format(ws, src_row=row - 1, dst_row=row, n_cols=6)
                ws.cell(row=row, column=COL_QUARTER).value = quarter
                ws.cell(row=row, column=COL_MES    ).value = mes
                ws.cell(row=row, column=COL_SEMANA ).value = wd.semana
                ws.cell(row=row, column=COL_WB     ).value = wd.whole_bird
                ws.cell(row=row, column=COL_WOGS   ).value = wd.wogs
                ws.cell(row=row, column=COL_COMP   ).value = wd.composite

            existing.add(wd.semana)
            novos += 1
            log.append(
                f"  [{wd.semana}] {quarter} | "
                f"WB={wd.whole_bird:.2f}  WOGS={wd.wogs:.2f}  Comp={wd.composite:.2f}"
            )

        if novos > 0 and not dry_run:
            wb.save(str(path)); drive_io.update_file(fid, path)
            print(f"Salvo. {novos} semana(s) adicionada(s), {skipped} ja existia(m).")
        elif novos > 0:
            print(f"[DRY-RUN] {novos} semana(s) seriam adicionadas, {skipped} ja existem.")
        else:
            print(f"Nenhuma alteracao — {skipped} semana(s) ja na planilha.")

        for entry in log:
            print(entry)

        return {"novos": novos, "skipped": skipped}
    finally:
        wb.close()


# ─────────────────────────────────────────────────────────────────────────────
# 4. Main
# ─────────────────────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Atualiza aba NationalComposite no FrangoUS.xlsm com dados USDA AMS."
    )
    parser.add_argument(
        "--desde", metavar="ANO", type=int,
        help="Backfill historico desde o ano informado (ex: --desde 2022).",
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Mostra o que seria atualizado sem gravar na planilha.",
    )
    args = parser.parse_args()

    print(f"[{dt.datetime.now():%Y-%m-%d %H:%M}] atualiza_national_composite.py")
    print(f"Arquivo: {XLSX_NAME}  |  Aba: {SHEET_NAME}")

    if args.desde:
        print(f"\nModo: backfill desde {args.desde}")
        week_list = fetch_historical(start_year=args.desde)
        print(f"Total: {len(week_list)} semana(s) encontrada(s).")
    else:
        print("\nModo: semana atual")
        week_list = fetch_recent()

    print(f"\nAtualizando {XLSX_NAME}...")
    result = update_sheet(week_list, dry_run=args.dry_run)

    print(
        f"\n[{dt.datetime.now():%Y-%m-%d %H:%M}] Concluido — "
        f"{result['novos']} nova(s), {result['skipped']} ja existia(m)."
    )


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        raise SystemExit(130)
