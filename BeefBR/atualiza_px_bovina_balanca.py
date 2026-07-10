"""
atualiza_px_bovina_balanca.py

Atualiza precos de exportacao (bovina e frango) em BeefBR.xlsm e FrangoBR.xlsm
com dados da Balanca Comercial preliminar (MDIC).

Logica por coluna de preco (E em BeefBR, M em FrangoBR):
  - Meses ate o ultimo disponivel na aba SECEX -> formula =SECEX!XX (verde)
  - Mes de referencia da balanca               -> valor da balanca / 1000 (azul)

Logica por coluna de formula (J em BeefBR, N em FrangoBR):
  - Sempre =preco * cambio (laranja), para todos os meses preenchidos acima

Nota sobre SECEX no FrangoBR:
  A aba SECEX fica dentro de BeefBR.xlsm. Ao escrever formulas no FrangoBR,
  o codigo mantem o BeefBR aberto na mesma instancia do Excel para que a
  referencia cruzada [BeefBR.xlsm]SECEX!X{linha} seja resolvida corretamente.
  Se o FrangoBR ja tiver uma aba SECEX propria, troque FRAN_SECEX_PREFIX
  de "[BeefBR.xlsm]SECEX!" para "SECEX!".
"""

import sys
import re
import io
import datetime
import requests
from pathlib import Path
from bs4 import BeautifulSoup
import openpyxl

# ── Paths ─────────────────────────────────────────────────────────────────────
DRIVE_ROOT = Path(
    r"G:\.shortcut-targets-by-id\11q1ngdqbySnDMCQSP4fiwPztT6ju2F4Z"
)
BEEFBR_NAME   = "BeefBR.xlsm"
FRANGOBR_NAME = "FrangoBR.xlsm"


def localizar_planilha(nome: str) -> Path:
    """Localiza uma planilha sem depender dos nomes das pastas intermediárias."""
    encontrados = list(DRIVE_ROOT.rglob(nome))

    if not encontrados:
        raise FileNotFoundError(
            f"Planilha {nome!r} não encontrada dentro de {DRIVE_ROOT}"
        )

    if len(encontrados) > 1:
        caminhos = "\n".join(f"  - {path}" for path in encontrados)
        raise RuntimeError(
            f"Foram encontradas várias planilhas chamadas {nome!r}:\n{caminhos}"
        )

    return encontrados[0]

# ── URL balanca ───────────────────────────────────────────────────────────────
BALANCA_URL = "https://balanca.economia.gov.br/balanca/pg_principal_bc/principais_resultados.html"

# ── Produtos na planilha da balanca ──────────────────────────────────────────
PRODUTO_BOVI = "Carne bovina fresca, refrigerada ou congelada"
PRODUTO_FRAN = "Carnes de aves e suas miudezas"   # prefixo da linha

# ── Cores (Excel BGR como inteiro) ───────────────────────────────────────────
# BLUE   = #0070C0  -> R=0,   G=112, B=192 -> 0 + 112*256 + 192*65536
# ORANGE = #ED7D31  -> R=237, G=125, B=49  -> 237 + 125*256 + 49*65536
# GREEN  = #00B050  -> R=0,   G=176, B=80  -> 0 + 176*256 + 80*65536 (fallback)
BLUE           = 12611584
ORANGE         = 3243501
GREEN_FALLBACK = 5287936

# ── Configuracao das abas ─────────────────────────────────────────────────────
#  col_price   : coluna do preco USD (E=5 em BeefBR, M=13 em FrangoBR)
#  col_formula : coluna da formula preco*cambio (J=10 em BeefBR, N=14 em FrangoBR)
#  col_fx      : coluna do cambio (AN=40 em BeefBR, T=20 em FrangoBR)
#  secex_col   : coluna na aba SECEX (W=23 bovina, X=24 frango)
#  secex_prefix: prefixo da formula SECEX (mesma pasta ou referencia cruzada)
#  sheet_name  : nome da aba na planilha destino
BLACK = 0  # preto / cor automatica no Excel

BEEF_CFG = dict(
    col_price          = 5,      # E
    col_formula        = 10,     # J
    col_fx             = 40,     # AN
    formula_past_color = None, # meses passados col J: verde (mesmo das celulas SECEX)
    sheet_name         = "BeefBR",
    label              = "BeefBR",
)
FRAN_CFG = dict(
    col_price          = 13,     # M
    col_formula        = 14,     # N
    col_fx             = 20,     # T
    formula_past_color = BLACK,  # meses passados col N: preto
    sheet_name         = "FrangoBR",
    label              = "FrangoBR",
)

# ── Constantes Excel ──────────────────────────────────────────────────────────
XL_CALC_MANUAL = -4135
XL_CALC_AUTO   = -4105

MONTH_PT = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
            "Jul", "Ago", "Set", "Out", "Nov", "Dez"]

MON_MAP = {
    "jan": 1, "fev": 2, "mar": 3, "abr": 4,
    "mai": 5, "jun": 6, "jul": 7, "ago": 8,
    "set": 9, "out": 10, "nov": 11, "dez": 12,
}


# ── Helpers de linha/coluna ───────────────────────────────────────────────────

def col_letter(n: int) -> str:
    """Numero de coluna (1-based) para letra Excel."""
    result = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def sheet_row(year: int, month: int) -> int:
    """Linha na aba (BeefBR ou FrangoBR). Linha 5 = jan/2000."""
    return 5 + (year - 2000) * 12 + (month - 1)


def row_to_yearmonth(row: int) -> tuple[int, int]:
    """Inverso de sheet_row."""
    offset = row - 5
    return 2000 + offset // 12, 1 + offset % 12


# ── Download ──────────────────────────────────────────────────────────────────

def download_setores_produtos() -> bytes:
    """Acessa o site da balanca comercial e baixa o Setores_Produtos.xlsx."""
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    session = requests.Session()
    session.verify = False
    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        )
    })

    print(f"  Acessando: {BALANCA_URL}")
    r = session.get(BALANCA_URL, timeout=30)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")

    link = None
    for a in soup.find_all("a", href=True):
        if re.search(r"[Ss]etores.?[Pp]rodutos", a["href"]):
            link = a["href"]
            break
    if not link:
        for a in soup.find_all("a", href=True):
            href = a["href"].lower()
            txt  = a.get_text(strip=True).lower()
            if ".xlsx" in href and ("setor" in href or "produto" in href):
                link = a["href"]
                break
            if ".xlsx" in href and ("setor" in txt or "produto" in txt):
                link = a["href"]
                break
    if not link:
        for a in soup.find_all("a", href=True):
            if a["href"].lower().endswith(".xlsx"):
                link = a["href"]
                break
    if not link:
        raise ValueError("Link de download (.xlsx) nao encontrado na pagina")

    if not link.startswith("http"):
        from urllib.parse import urljoin
        link = urljoin(BALANCA_URL, link)

    print(f"  Baixando: {link}")
    r2 = session.get(link, timeout=60)
    r2.raise_for_status()
    return r2.content


# ── Correcao do xlsx ──────────────────────────────────────────────────────────

def fix_missing_drawings(content: bytes) -> bytes:
    """Remove referencias a drawings inexistentes que quebram o openpyxl."""
    import zipfile
    import xml.etree.ElementTree as ET

    NS = "http://schemas.openxmlformats.org/package/2006/relationships"
    ET.register_namespace("", NS)

    buf_in  = io.BytesIO(content)
    buf_out = io.BytesIO()

    with zipfile.ZipFile(buf_in, "r") as zin:
        existing = set(zin.namelist())
        with zipfile.ZipFile(buf_out, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename.endswith(".rels"):
                    try:
                        root = ET.fromstring(data)
                        rels_dir = "/".join(
                            info.filename.replace("_rels/", "").split("/")[:-1]
                        )
                        to_remove = []
                        for rel in root.findall(f"{{{NS}}}Relationship"):
                            target = rel.get("Target", "")
                            if "drawing" not in target.lower():
                                continue
                            if target.startswith("/"):
                                resolved = target.lstrip("/")
                            else:
                                parts = (rels_dir + "/" + target).split("/")
                                norm: list[str] = []
                                for p in parts:
                                    if p == "..":
                                        if norm:
                                            norm.pop()
                                    elif p and p != ".":
                                        norm.append(p)
                                resolved = "/".join(norm)
                            if resolved not in existing:
                                to_remove.append(rel)
                        if to_remove:
                            for rel in to_remove:
                                root.remove(rel)
                            data = ET.tostring(root, encoding="unicode").encode("utf-8")
                    except Exception:
                        pass
                zout.writestr(info, data)

    buf_out.seek(0)
    return buf_out.read()


# ── Parse do xlsx ─────────────────────────────────────────────────────────────

def _try_parse_month_str(val: str) -> tuple[int, int] | None:
    val = val.strip().lower()
    m = re.match(r"([a-záéíóúâêôãõü]+)[/\-\s](\d{2,4})$", val)
    if not m:
        return None
    mon_str = m.group(1)[:3]
    yr_str  = m.group(2)
    if mon_str not in MON_MAP:
        return None
    yr = int(yr_str)
    if yr < 100:
        yr += 2000
    return yr, MON_MAP[mon_str]


def parse_data_month(ws) -> tuple[int, int]:
    """Detecta o mes/ano dos dados na coluna J do xlsx."""
    for row in ws.iter_rows(min_row=1, max_row=10, max_col=11):
        if len(row) < 10:
            continue
        cell = row[9]
        val  = cell.value
        if val is None:
            continue
        if hasattr(val, "year") and hasattr(val, "month"):
            return val.year, val.month
        if isinstance(val, str):
            result = _try_parse_month_str(val)
            if result:
                return result
        if isinstance(val, (int, float)) and 30000 < val < 60000:
            try:
                d = datetime.date(1899, 12, 30) + datetime.timedelta(days=int(val))
                return d.year, d.month
            except Exception:
                pass
    raise ValueError("Nao foi possivel determinar o mes dos dados na coluna J")


def parse_prices(content: bytes) -> tuple[dict[str, float], int, int]:
    """
    Le o xlsx e retorna ({produto_key: price_raw}, year, month).
    price_raw e o valor bruto da coluna J (antes de /1000).
    """
    content = fix_missing_drawings(content)
    wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)

    ws = None
    for name in wb.sheetnames:
        if re.search(r"setor|produto", name, re.I):
            ws = wb[name]
            break
    if ws is None:
        ws = wb.active

    year, month = parse_data_month(ws)

    targets = {
        "bovi": PRODUTO_BOVI.lower(),
        "fran": PRODUTO_FRAN.lower(),
    }
    prices: dict[str, float] = {}

    for row in ws.iter_rows():
        if len(row) < 10:
            continue
        cell_a = row[0]
        if not (cell_a.value and isinstance(cell_a.value, str)):
            continue
        cell_val = cell_a.value.strip().lower()
        for key, produto in targets.items():
            if key not in prices and produto in cell_val:
                raw = row[9].value
                if raw is not None:
                    try:
                        prices[key] = float(str(raw).replace(",", "."))
                    except ValueError:
                        pass

    for key in targets:
        if key not in prices:
            raise ValueError(f"Produto '{key}' nao encontrado na planilha da balanca")

    print(f"  Bovina: {prices['bovi']:,.2f} -> {prices['bovi']/1000:.4f} USD/unid")
    print(f"  Frango: {prices['fran']:,.2f} -> {prices['fran']/1000:.4f} USD/unid")
    print(f"  Referencia: {MONTH_PT[month-1]}/{year}")
    return prices, year, month


# ── SECEX helpers ─────────────────────────────────────────────────────────────

def build_secex_map(ws_secex) -> dict[tuple[int, int], int]:
    """Constroi {(year, month): secex_row} a partir da aba SECEX."""
    date_to_row: dict[tuple[int, int], int] = {}
    for r in range(3, 5000):
        val = ws_secex.Cells(r, 2).Value
        if val is None:
            break
        try:
            date_to_row[(val.year, val.month)] = r
        except AttributeError:
            continue
    return date_to_row


def find_secex_ref(ws, col_price: int) -> tuple[str, int, int] | None:
    """
    Lê uma célula verde já existente em col_price e extrai da fórmula:
      - o prefixo   ex: 'SECEX!' ou '[BeefBR.xlsm]SECEX!'
      - a letra     ex: 'F'
      - o numero    ex: 6
    Retorna (prefixo, col_numero, cor_fonte) ou None se nao encontrar.

    Isso evita hardcodar qual coluna do SECEX usar: o codigo se adapta
    automaticamente ao que ja esta na planilha (colunas C->M = Brasil).
    """
    try:
        XL_UP  = -4162
        bottom = ws.Cells(ws.Rows.Count, col_price).End(XL_UP).Row
        for r in range(bottom, 4, -1):
            cell = ws.Cells(r, col_price)
            if not cell.HasFormula:
                continue
            formula = str(cell.Formula)
            if "SECEX" not in formula.upper():
                continue
            m = re.search(r"((?:\[[^\]]+\])?SECEX!)([A-Z]+)\d+", formula, re.I)
            if m:
                prefix  = m.group(1)
                col_str = m.group(2).upper()
                col_num = 0
                for ch in col_str:
                    col_num = col_num * 26 + (ord(ch) - ord("A") + 1)
                color = int(cell.Font.Color)
                return prefix, col_num, color
    except Exception:
        pass
    return None


def find_last_secex_month(
    ws_secex,
    secex_col: int,
    date_to_row: dict[tuple[int, int], int],
) -> tuple[int, int] | None:
    """
    Retorna o ultimo (year, month) cujo valor na coluna secex_col seja > 0.
    Sem restricao de mes: o SECEX tem prioridade sobre a balanca em qualquer mes.
    """
    last_ym = None
    for ym in sorted(date_to_row):
        r = date_to_row[ym]
        v = ws_secex.Cells(r, secex_col).Value
        try:
            if v is not None and float(v) > 0:
                last_ym = ym
        except (TypeError, ValueError):
            pass
    return last_ym


def secex_has_value(ws_secex, secex_col: int, secex_row: int) -> bool:
    """Retorna True se a celula (secex_row, secex_col) tem valor > 0."""
    try:
        v = ws_secex.Cells(secex_row, secex_col).Value
        return v is not None and float(v) > 0
    except (TypeError, ValueError):
        return False


# ── Escrita na aba ────────────────────────────────────────────────────────────

def _copy_number_format(ws, src_row: int, dst_row: int, col: int) -> None:
    """Copia apenas o NumberFormat da linha acima (sem cor)."""
    try:
        ws.Cells(dst_row, col).NumberFormat = ws.Cells(src_row, col).NumberFormat
    except Exception:
        pass


def update_price_sheet(
    ws,
    ws_secex,
    date_to_row: dict[tuple[int, int], int],
    price_raw: float,
    cfg: dict,
    data_year: int,
    data_month: int,
    excel,
) -> None:
    """
    Atualiza a coluna de preco e a coluna de formula em uma aba (BeefBR ou FrangoBR).

    Para meses com dado na aba SECEX:
      col_price   : formula =SECEX!XX (verde)
      col_formula : formula =SECEX!KK (verde, BeefBR) ou =price*fx (preto, FrangoBR)

    Para o mes de referencia da balanca (data_year/data_month):
      - Se SECEX ja tiver dado para esse mes -> nao faz nada (prioridade SECEX)
      - Caso contrario:
          col_price   : valor da balanca / 1000 (azul)
          col_formula : =price*fx (laranja BeefBR, preto FrangoBR)
    """
    col_price   = cfg["col_price"]
    col_formula = cfg["col_formula"]
    col_fx      = cfg["col_fx"]
    label       = cfg["label"]

    price_l = col_letter(col_price)
    fx_l    = col_letter(col_fx)

    # formula_past_color cfg: None = verde SECEX (BeefBR J), BLACK = preto (FrangoBR N)
    _past_color_cfg = cfg.get("formula_past_color", ORANGE)

    # ── 1. Detectar referencia SECEX para col_price (E / M) ──────────────────
    secex_ref_price = find_secex_ref(ws, col_price)
    if secex_ref_price is None:
        print(f"  {label}: AVISO - nenhuma formula SECEX em col {col_price}. Pulando meses passados.")
        secex_prefix = None
        secex_col    = None
        green        = GREEN_FALLBACK
    else:
        secex_prefix, secex_col, green = secex_ref_price
        green = int(green)
        print(f"  {label}: SECEX preco -> {secex_prefix}{col_letter(secex_col)}N "
              f"(col {secex_col}, cor 0x{green:06X})")

    # ── 2. Detectar referencia SECEX para col_formula (J / N) ────────────────
    secex_ref_formula = find_secex_ref(ws, col_formula)
    if secex_ref_formula is not None:
        sec_prefix_f, secex_col_f, green_f = secex_ref_formula
        formula_past_color = int(green_f)
        print(f"  {label}: SECEX formula -> {sec_prefix_f}{col_letter(secex_col_f)}N "
              f"(col {secex_col_f}, cor 0x{formula_past_color:06X})")
    else:
        sec_prefix_f       = None
        secex_col_f        = None
        formula_past_color = GREEN_FALLBACK if _past_color_cfg is None else _past_color_cfg

    data_ym  = (data_year, data_month)
    row_data = sheet_row(data_year, data_month)

    # ── 3. Preencher meses com dado no SECEX ─────────────────────────────────
    secex_rows_written: list[int] = []
    secex_has_data_month          = False

    if secex_col is not None:
        last_secex_ym = find_last_secex_month(ws_secex, secex_col, date_to_row)

        if last_secex_ym:
            last_yr, last_mo = last_secex_ym
            row_last_secex   = sheet_row(last_yr, last_mo)

            # Ultima linha que ja tem formula SECEX nessa aba
            last_secex_formula_r = 4
            for r in range(row_last_secex, 4, -1):
                try:
                    cell = ws.Cells(r, col_price)
                    if cell.HasFormula and "SECEX" in str(cell.Formula).upper():
                        last_secex_formula_r = r
                        break
                except Exception:
                    pass

            # Preencher do proximo ao ultimo SECEX disponivel
            for r in range(last_secex_formula_r + 1, row_last_secex + 1):
                yr, mo  = row_to_yearmonth(r)
                secex_r = date_to_row.get((yr, mo))
                if secex_r is None:
                    continue

                # col_price: formula SECEX (verde)
                _copy_number_format(ws, r - 1, r, col_price)
                ws.Cells(r, col_price).Formula    = f"={secex_prefix}{col_letter(secex_col)}{secex_r}"
                ws.Cells(r, col_price).Font.Color = green

                # col_formula: formula SECEX se detectada, senao price*fx
                _copy_number_format(ws, r - 1, r, col_formula)
                if sec_prefix_f is not None:
                    ws.Cells(r, col_formula).Formula = (
                        f"={sec_prefix_f}{col_letter(secex_col_f)}{secex_r}"
                    )
                else:
                    ws.Cells(r, col_formula).Formula = f"={price_l}{r}*{fx_l}{r}"
                ws.Cells(r, col_formula).Font.Color = formula_past_color

                secex_rows_written.append(r)

            if secex_rows_written:
                print(f"  {label}: {len(secex_rows_written)} mes(es) SECEX preenchidos "
                      f"(ate {MONTH_PT[last_mo-1]}/{last_yr})")
            else:
                print(f"  {label}: SECEX ja atualizado ate {MONTH_PT[last_mo-1]}/{last_yr}")

            # Verificar se SECEX cobre o mes de referencia da balanca
            data_secex_row = date_to_row.get(data_ym)
            secex_has_data_month = (
                data_secex_row is not None
                and secex_has_value(ws_secex, secex_col, data_secex_row)
            )

    # ── 4. Mes de referencia da balanca (somente se SECEX nao cobrir) ─────────
    if secex_has_data_month:
        print(f"  {label}: SECEX ja cobre {MONTH_PT[data_month-1]}/{data_year} - balanca ignorada")
    else:
        _copy_number_format(ws, row_data - 1, row_data, col_price)
        ws.Cells(row_data, col_price).Value      = price_raw / 1000
        ws.Cells(row_data, col_price).Font.Color = BLUE
        print(f"  {label}: mes balanca ({MONTH_PT[data_month-1]}/{data_year}) = "
              f"{price_raw/1000:.4f} (azul)")

        formula_curr_color = ORANGE
        _copy_number_format(ws, row_data - 1, row_data, col_formula)
        ws.Cells(row_data, col_formula).Formula    = f"={price_l}{row_data}*{fx_l}{row_data}"
        ws.Cells(row_data, col_formula).Font.Color = formula_curr_color
        cor_curr = "laranja" if formula_curr_color == ORANGE else "preto"
        print(f"  {label}: col {col_letter(col_formula)} mes balanca ({cor_curr})")

    if secex_rows_written:
        cor_past = "verde" if secex_ref_formula is not None else (
            "preto" if formula_past_color == BLACK else "verde"
        )
        print(f"  {label}: {len(secex_rows_written)} mes(es) col {col_letter(col_formula)} "
              f"preenchidos ({cor_past})")


# ── Helpers Excel ────────────────────────────────────────────────────────────

# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    try:
        import win32com.client
    except ImportError:
        print("ERRO: pywin32 nao instalado. Execute: pip install pywin32")
        sys.exit(1)

    try:
        beefbr_path = localizar_planilha(BEEFBR_NAME)
        frangobr_path = localizar_planilha(FRANGOBR_NAME)
    except (FileNotFoundError, RuntimeError) as exc:
        print(f"ERRO: {exc}")
        sys.exit(1)

    today = datetime.date.today()
    curr_year, curr_month = today.year, today.month
    print(f"Preco Exportacao (Balanca) -> BeefBR + FrangoBR  --  {today:%d/%m/%Y}\n")

    # ── Passo 1: Baixar e interpretar xlsx ────────────────────────────────────
    content = download_setores_produtos()
    prices, data_year, data_month = parse_prices(content)

    print()

    # ── Passo 2: Abrir Excel ──────────────────────────────────────────────────
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible        = False
    excel.DisplayAlerts  = False
    excel.ScreenUpdating = False

    try:
        # BeefBR fica aberto do inicio ao fim:
        #   - contem a aba SECEX referenciada pelo FrangoBR
        #   - evita fechar/reabrir (que gerava prompt de "mesmo nome")
        wb_beef  = excel.Workbooks.Open(str(beefbr_path))
        excel.Calculation = XL_CALC_MANUAL
        ws_beef  = wb_beef.Sheets("BeefBR")
        ws_secex = wb_beef.Sheets("SECEX")

        date_to_row = build_secex_map(ws_secex)
        print(f"  SECEX: {len(date_to_row)} meses mapeados\n")

        # ── BeefBR ────────────────────────────────────────────────────────────
        update_price_sheet(
            ws_beef, ws_secex, date_to_row,
            prices["bovi"], BEEF_CFG,
            data_year, data_month, excel,
        )
        excel.Calculation = XL_CALC_AUTO
        wb_beef.Save()
        print(f"\n  Salvo: BeefBR.xlsm")

        # ── FrangoBR ──────────────────────────────────────────────────────────
        # BeefBR permanece aberto -> [BeefBR.xlsm]SECEX! resolve sem prompt
        wb_fran = excel.Workbooks.Open(str(frangobr_path))
        excel.Calculation = XL_CALC_MANUAL
        ws_fran = wb_fran.Sheets("FrangoBR")

        print()
        update_price_sheet(
            ws_fran, ws_secex, date_to_row,
            prices["fran"], FRAN_CFG,
            data_year, data_month, excel,
        )
        excel.Calculation = XL_CALC_AUTO
        wb_fran.Save()
        wb_fran.Close(SaveChanges=False)
        print(f"\n  Salvo: FrangoBR.xlsm")

        wb_beef.Close(SaveChanges=False)

    except Exception:
        import traceback
        traceback.print_exc()
        sys.exit(1)
    finally:
        try:
            excel.ScreenUpdating = True
            excel.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
