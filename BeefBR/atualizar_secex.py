import requests
from bs4 import BeautifulSoup
from datetime import date
import openpyxl
import xlwings as xw
import sys
from pathlib import Path

# ─── CONFIGURAÇÃO ────────────────────────────────────────────────────────────
DRIVE_ROOT = Path(
    r"G:\.shortcut-targets-by-id\11q1ngdqbySnDMCQSP4fiwPztT6ju2F4Z"
)
XLSX_NAME = "BeefBR.xlsm"
SHEET_NAME = "SECEX"


def localizar_planilha() -> Path:
    """Localiza a planilha sem depender dos nomes das pastas intermediárias."""
    encontrados = list(DRIVE_ROOT.rglob(XLSX_NAME))

    if not encontrados:
        raise FileNotFoundError(
            f"Planilha {XLSX_NAME!r} não encontrada dentro de {DRIVE_ROOT}"
        )

    if len(encontrados) > 1:
        caminhos = "\n".join(f"  - {path}" for path in encontrados)
        raise RuntimeError(
            f"Foram encontradas várias planilhas chamadas {XLSX_NAME!r}:\n"
            f"{caminhos}"
        )

    return encontrados[0]


FILE_PATH = str(localizar_planilha())
# ─────────────────────────────────────────────────────────────────────────────

# ── PASSO 1: Ler estrutura com openpyxl (sem salvar) ─────────────────────────
print("Lendo estrutura da planilha...")
wb_opx = openpyxl.load_workbook(FILE_PATH)  # sem data_only para detectar fórmulas
ws_opx = wb_opx[SHEET_NAME]

date_to_row = {}
formula_cells = set()  # (row, col) com fórmula — nunca tocar

for row in ws_opx.iter_rows(min_row=3):
    cell_b = row[1]  # coluna B
    val = cell_b.value
    if val and hasattr(val, "year") and hasattr(val, "month"):
        date_to_row[(val.year, val.month)] = cell_b.row

# Mapear todas as células com fórmula nas colunas alvo
TARGET_COLS = [18, 19, 20, 22, 23, 24]
for (year, month), row_num in date_to_row.items():
    for col in TARGET_COLS:
        cell = ws_opx.cell(row=row_num, column=col)
        if cell.data_type == "f":
            formula_cells.add((row_num, col))

wb_opx.close()

today  = date.today()
cutoff = (today.year, today.month)

# Verificar lacunas (usando xlwings para ler valores reais)
print("Verificando lacunas via xlwings...")
app    = xw.App(visible=False)
wb_xl  = app.books.open(FILE_PATH)
ws_xl  = wb_xl.sheets[SHEET_NAME]

gap_months = set()
for (year, month), row_num in sorted(date_to_row.items()):
    if (year, month) >= cutoff:
        continue
    for col in TARGET_COLS:
        if (row_num, col) in formula_cells:
            continue
        if ws_xl.cells(row_num, col).value is None:
            gap_months.add((year, month))
            break

if not gap_months:
    last = max((ym for ym in date_to_row if ym < cutoff), default=None)
    wb_xl.close()
    app.quit()
    MONTH_PT = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]
    if last:
        print(f"Planilha já atualizada até {MONTH_PT[last[1]-1]}-{last[0]}. Nenhuma ação necessária.")
    else:
        print("Nenhuma data encontrada na planilha.")
    sys.exit(0)

print(f"Lacunas encontradas em {len(gap_months)} mês(es). Consultando USDA GATS...")

# ── PASSO 2: Consultar USDA GATS ─────────────────────────────────────────────
session = requests.Session()
session.headers.update({"User-Agent": "Mozilla/5.0"})
BASE = "https://apps.fas.usda.gov/GATS/"

session.get(BASE + "detectscreen.aspx?action=set&res=1920x1080&d=24&pw=1920")
r_home = session.get(BASE + "default.aspx")
soup_home = BeautifulSoup(r_home.text, "html.parser")

# Navegar até a página de Standard Query (onde fica o formulário completo)
std_link = soup_home.find("a", string=lambda t: t and "Standard Query" in t)
query_url = (BASE + std_link["href"].lstrip("/")) if std_link else (BASE + "ExpressQuery1.aspx")
print(f"  URL Standard Query: {query_url}")

r0    = session.get(query_url)
soup0 = BeautifulSoup(r0.text, "html.parser")

# Identificar URL de POST do formulário
form = soup0.find("form")
form_action = form.get("action", "") if form else ""
if form_action:
    post_url = BASE + form_action.lstrip("/") if not form_action.startswith("http") else form_action
else:
    post_url = query_url
print(f"  URL de POST: {post_url}")


def get_form_fields(soup):
    fields = {}
    form = soup.find("form") or soup
    for inp in form.find_all("input"):
        name = inp.get("name")
        if not name:
            continue
        t = inp.get("type", "text").lower()
        if t in ("hidden", "text", "password"):
            fields[name] = inp.get("value", "")
        elif t == "checkbox" and inp.get("checked"):
            fields[name] = inp.get("value", "on")
        elif t == "radio" and inp.get("checked"):
            fields[name] = inp.get("value", "")
    for sel in form.find_all("select"):
        name = sel.get("name")
        if not name:
            continue
        selected = sel.find("option", selected=True) or sel.find("option")
        fields[name] = selected.get("value", "") if selected else ""
    return fields


PRODUCTS = [
    ("suina",  "FATUS", "1200AG"),
    ("bovina", "FATUS", "1130AG"),
    ("frango", "FAS",   "0504AT"),
]

month_data = {p[0]: {} for p in PRODUCTS}

PFX = "ctl00$ContentPlaceHolder1$"

def parse_gats_table(table):
    """39 células por linha de dados:
    cells[9]='YYYY-YYYY', cells[10]='MT',
    cells[11]=Jan UV, cells[12]=Jan Qty, cells[13]=Fev UV, ..., cells[34]=Dez Qty"""
    results = {}
    for tr in table.find_all("tr"):
        cells = tr.find_all("td")
        if len(cells) != 39:
            continue
        year_text = cells[9].get_text(strip=True)
        if "-" not in year_text or not year_text[:4].isdigit():
            continue
        yr = int(year_text[:4])
        for mo in range(1, 13):
            uv_raw  = cells[11 + (mo - 1) * 2].get_text(strip=True).replace(",", "")
            qty_raw = cells[12 + (mo - 1) * 2].get_text(strip=True).replace(",", "")
            if uv_raw in ("(D)", "(X)", "") or qty_raw in ("(D)", "(X)", ""):
                continue
            try:
                results[(yr, mo)] = {
                    "uv_usd_per_mt": float(uv_raw),
                    "qty_mt":        float(qty_raw),
                }
            except ValueError:
                continue
    return results

for prod_name, grupo, produto in PRODUCTS:
    print(f"  Consultando {prod_name}...")

    # POST 1: atualiza ViewState para o grupo de produto correto
    payload1 = get_form_fields(soup0)
    payload1["__EVENTTARGET"]         = PFX + "ddlProductGroup"
    payload1["__EVENTARGUMENT"]       = ""
    payload1[PFX + "ddlProductGroup"] = grupo

    r1    = session.post(post_url, data=payload1)
    soup1 = BeautifulSoup(r1.text, "html.parser")

    # POST 2: UV + Qty combinados em uma única consulta
    payload2 = get_form_fields(soup1)
    payload2["__EVENTTARGET"]         = ""
    payload2["__EVENTARGUMENT"]       = ""
    payload2[PFX + "ddlProductGroup"] = grupo
    payload2[PFX + "lb_Products"]     = produto
    payload2[PFX + "lb_Partners"]     = "R00"
    payload2[PFX + "ddlDateSeries"]   = "Monthly"
    payload2[PFX + "ddlStartYear"]    = "2000"
    payload2[PFX + "ddlEndYear"]      = str(today.year)
    payload2[PFX + "ddlValueType"]    = "UVAL"
    payload2[PFX + "ddlValueUnit"]    = "D"
    payload2[PFX + "btnRetrieveData"] = "Retrieve Data"

    r2    = session.post(post_url, data=payload2)
    soup2 = BeautifulSoup(r2.text, "html.parser")
    table = soup2.find("table", id=lambda x: x and "GridView" in x)

    if not table:
        print(f"    AVISO: tabela não encontrada para {prod_name}")
        continue

    records = parse_gats_table(table)
    for (yr, mo), vals in records.items():
        month_data[prod_name][(yr, mo)] = vals
    print(f"    {len(records)} meses obtidos")

# ── PASSO 3: Preencher lacunas com xlwings ────────────────────────────────────
COL_MAP = {
    18: ("suina",  "qty_mt"),
    19: ("bovina", "qty_mt"),
    20: ("frango", "qty_mt"),
    22: ("suina",  "uv_usd_per_mt"),
    23: ("bovina", "uv_usd_per_mt"),
    24: ("frango", "uv_usd_per_mt"),
}

# Ler formatação da última célula já preenchida em cada coluna alvo
col_fmt = {}
for col in COL_MAP:
    for (year, month), row_num in sorted(date_to_row.items(), reverse=True):
        if (year, month) >= cutoff:
            continue
        if (row_num, col) in formula_cells:
            continue
        cell = ws_xl.cells(row_num, col)
        if cell.value is not None:
            a = cell.api
            col_fmt[col] = {
                "font_name":   a.Font.Name,
                "font_size":   a.Font.Size,
                "font_bold":   a.Font.Bold,
                "font_color":  a.Font.Color,
                "font_italic": a.Font.Italic,
                "h_align":     a.HorizontalAlignment,
                "v_align":     a.VerticalAlignment,
                "number_fmt":  a.NumberFormat,
                "wrap_text":   a.WrapText,
            }
            break

filled_months = set()

for (year, month), row_num in sorted(date_to_row.items()):
    if (year, month) >= cutoff:
        continue
    for col, (prod, field) in COL_MAP.items():
        if (row_num, col) in formula_cells:
            continue

        cell    = ws_xl.cells(row_num, col)
        entry   = month_data.get(prod, {}).get((year, month), {})
        raw_val = entry.get(field)

        new_val = round(raw_val / 1000, 4 if field == "qty_mt" else 6) if raw_val is not None else None

        if cell.value is None and new_val is not None:
            cell.value = new_val
            filled_months.add((year, month))
            if col in col_fmt:
                f   = col_fmt[col]
                api = cell.api
                api.Font.Name           = f["font_name"]
                api.Font.Size           = f["font_size"]
                api.Font.Bold           = f["font_bold"]
                api.Font.Color          = f["font_color"]
                api.Font.Italic         = f["font_italic"]
                api.HorizontalAlignment = f["h_align"]
                api.VerticalAlignment   = f["v_align"]
                api.NumberFormat        = f["number_fmt"]
                api.WrapText            = f["wrap_text"]

wb_xl.save()
wb_xl.close()
app.quit()
print("Planilha salva.")

# ── PASSO 4: Relatório ────────────────────────────────────────────────────────
MONTH_PT = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]

if filled_months:
    sorted_m   = sorted(filled_months)
    months_str = ", ".join(f"{MONTH_PT[m-1]}-{y}" for y, m in sorted_m)
    print(f"{len(sorted_m)} mês(es) preenchido(s): {months_str}")
else:
    print("Nenhum novo mês preenchido (dados podem não estar disponíveis no GATS ainda).")

# "último completo" = meses sem lacuna original + o que foi preenchido agora
already_complete = {ym for ym in date_to_row if ym < cutoff} - gap_months
last = max(already_complete | filled_months, default=None)
if last:
    print(f"Último mês completo na planilha: {MONTH_PT[last[1]-1]}-{last[0]}")
