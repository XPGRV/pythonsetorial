"""
Atualização mensal ERS USDA — aba 'Abates' de BeefUS.xlsm

Fonte : ERS USDA – meat-statistics-tables-historical.xlsx
        https://www.ers.usda.gov/media/5529/meat-statistics-tables-historical.xlsx

Colunas de contagem escritas (Commercial, 1.000 cabeças):
  C = Bois - Cattle Slaughter (Steers)
  D = Novilhas - Heifers
  E = Vacas - Beef Cows
  F = Vacas Leiteiras - Dairy Cows
  G = Touros e Cervos - Bulls and Stags
  I = Bezerros - Calves

Colunas de fórmula (propagadas para novas linhas, não sobrescritas):
  A = Quarter (fórmula já existe em todas as linhas pré-populadas)
  H = Total Cattle  (=G+F+E+D+C)
  J:N = % por categoria

Colunas de peso escritas (Commercial avg live weight, libras por cabeça):
  P = Cattle
  Q = Calves
"""

import io
import time
import datetime as dt
from pathlib import Path

import requests
import openpyxl
import xlwings as xw


# ============================================================
# CONFIG
# ============================================================
DRIVE_ROOT = Path(
    r"G:\.shortcut-targets-by-id\11q1ngdqbySnDMCQSP4fiwPztT6ju2F4Z"
)
XLSX_NAME  = "BeefUS.xlsm"
SHEET_NAME = "Abates"
ANO_INICIO = 2000   # linha 3 da aba = jan/2000

USDA_URL    = "https://www.ers.usda.gov/media/5529/meat-statistics-tables-historical.xlsx"
TIMEOUT     = 120
MAX_RETRIES = 3
RETRY_SLEEP = 5

REVISION_MONTHS = 12   # meses para verificar revisões retroativamente

# Índices (0-based) na sheet SlaughterCounts-Full do USDA
SC_STEERS   = 2
SC_HEIFERS  = 3
SC_BEEFCOW  = 5
SC_DAIRYCOW = 6
SC_BULLS    = 7
SC_CALVES_C = 8

# Índices (0-based) na sheet SlaughterWeights-Full do USDA
SW_CATTLE = 1
SW_CALVES = 2

# Colunas de dados na planilha (letras Excel)
COL_STEERS    = "C"   # índice 3
COL_HEIFERS   = "D"
COL_BEEFCOW   = "E"
COL_DAIRYCOW  = "F"
COL_BULLS     = "G"
# col H = Total (fórmula)
COL_CALVES    = "I"   # índice 9  ← movido de H para I
COL_WT_CATTLE = "P"
COL_WT_CALVES = "Q"

# Colunas de fórmula propagadas para novas linhas
FORMULA_COL_TOTAL = "H"       # =G+F+E+D+C
FORMULA_RANGE_PCT = "J:N"     # %s

MONTH_ABBR = {
    "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
    "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12,
}


# ============================================================
# HELPERS
# ============================================================
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


def row_of(ano: int, mes: int) -> int:
    """Linha 3 = jan/2000."""
    return 3 + (ano - ANO_INICIO) * 12 + (mes - 1)


def parse_usda_date(s):
    """
    'Feb-2026' → (2026, 2).
    Retorna None para YTD ('Jan-Feb 2026') ou inválido.
    """
    if s is None:
        return None
    s = str(s).strip()
    parts = s.split("-")
    if len(parts) != 2:
        return None
    mon, yr = parts[0].strip(), parts[1].strip()
    if mon not in MONTH_ABBR:
        return None
    try:
        return int(yr), MONTH_ABBR[mon]
    except ValueError:
        return None


def _diff(a, b) -> bool:
    try:
        return abs(float(a) - float(b)) > 0.01
    except (TypeError, ValueError):
        return a != b


def _col_idx(letter: str) -> int:
    """'C' → 3, 'I' → 9, etc."""
    from openpyxl.utils import column_index_from_string
    return column_index_from_string(letter)


# ============================================================
# DOWNLOAD
# ============================================================
def download_usda() -> bytes:
    last_err = None
    for tentativa in range(1, MAX_RETRIES + 1):
        try:
            print(f"  Baixando USDA Excel (tentativa {tentativa}/{MAX_RETRIES})...")
            r = requests.get(USDA_URL, timeout=TIMEOUT)
            r.raise_for_status()
            print(f"  Download OK — {len(r.content) / 1024:.0f} KB")
            return r.content
        except Exception as e:
            last_err = e
            if tentativa < MAX_RETRIES:
                time.sleep(RETRY_SLEEP)
    raise RuntimeError(f"Download USDA falhou em {MAX_RETRIES} tentativas: {last_err}")


# ============================================================
# PARSE USDA
# ============================================================
def parse_usda(conteudo: bytes) -> dict:
    """
    Retorna:
      { (ano, mes): {
            'steers', 'heifers', 'beef_cows', 'dairy_cows',
            'bulls', 'calves', 'wt_cattle', 'wt_calves'
        }
      }
    """
    wb = openpyxl.load_workbook(io.BytesIO(conteudo), read_only=True, data_only=True)

    counts: dict[tuple, dict] = {}
    ws_c = wb["SlaughterCounts-Full"]
    for i, row in enumerate(ws_c.iter_rows(values_only=True)):
        if i < 3:
            continue
        parsed = parse_usda_date(row[0])
        if parsed is None:
            continue
        ano, mes = parsed
        if ano < ANO_INICIO:
            continue
        counts[(ano, mes)] = {
            "steers":     row[SC_STEERS],
            "heifers":    row[SC_HEIFERS],
            "beef_cows":  row[SC_BEEFCOW],
            "dairy_cows": row[SC_DAIRYCOW],
            "bulls":      row[SC_BULLS],
            "calves":     row[SC_CALVES_C],
            "wt_cattle":  None,
            "wt_calves":  None,
        }

    ws_w = wb["SlaughterWeights-Full"]
    for i, row in enumerate(ws_w.iter_rows(values_only=True)):
        if i < 3:
            continue
        parsed = parse_usda_date(row[0])
        if parsed is None:
            continue
        ano, mes = parsed
        if (ano, mes) in counts:
            counts[(ano, mes)]["wt_cattle"] = row[SW_CATTLE]
            counts[(ano, mes)]["wt_calves"] = row[SW_CALVES]

    print(f"  USDA: {len(counts)} meses mensais lidos ({ANO_INICIO}+)")
    return counts


# ============================================================
# MAIN
# ============================================================
def main():
    path = localizar_planilha()

    print(f"[{dt.datetime.now():%Y-%m-%d %H:%M}] Iniciando atualização USDA — BeefUS Abates...")

    print("\n[1/3] Baixando e processando Excel do USDA...")
    conteudo = download_usda()
    dados    = parse_usda(conteudo)

    if not dados:
        print("Nenhum dado encontrado. Encerrando.")
        return

    print(f"\n[2/3] Abrindo {path.name} via xlwings...")
    app = xw.App(visible=False, add_book=False)
    try:
        app.api.AskToUpdateLinks = False
        app.display_alerts       = False
        app.screen_updating      = False

        wb_xw = app.books.open(str(path), update_links=0, read_only=False)
        try:
            print("\n[3/3] Gravando dados na planilha...")
            resultado = processar(wb_xw, dados)
            if resultado["salvar"]:
                wb_xw.api.SaveLinkValues = True
                wb_xw.save()
                print("Arquivo salvo.")
            else:
                print("Nada alterado — arquivo NÃO salvo.")
        finally:
            wb_xw.close()
    finally:
        app.quit()

    print_relatorio(resultado)


# ============================================================
# PROCESSAMENTO
# ============================================================
def processar(wb_xw, dados: dict) -> dict:
    sh = wb_xw.sheets[SHEET_NAME]

    hoje         = dt.date.today()
    corte_rev    = hoje - dt.timedelta(days=REVISION_MONTHS * 31)
    corte_anoMes = (corte_rev.year, corte_rev.month)

    last_row = sh.used_range.last_cell.row

    # Lê col B (datas)
    datas_col_b = sh.range(f"B3:B{last_row}").value
    if not isinstance(datas_col_b, list):
        datas_col_b = [datas_col_b]

    # Lê contagens: C:G (5 cols) e I separado (para não incluir H=Total fórmula)
    def _read_matrix(rng_str, ncols):
        val = sh.range(rng_str).value
        if val is None:
            return []
        if not isinstance(val, list):
            return [[val] + [None] * (ncols - 1)]
        if not isinstance(val[0], list):
            return [[v] + [None] * (ncols - 1) for v in val]
        return val

    cg_mat  = _read_matrix(f"C3:G{last_row}", 5)   # steers, heifers, beef_cows, dairy_cows, bulls
    cal_col = sh.range(f"I3:I{last_row}").value     # calves (col I)
    if not isinstance(cal_col, list):
        cal_col = [cal_col]
    cal_col = [v[0] if isinstance(v, (list, tuple)) else v for v in cal_col]

    wt_mat  = _read_matrix(f"P3:Q{last_row}", 2)   # wt_cattle, wt_calves

    # Template = última linha com dado em C (para copiar formatos/fórmulas)
    template_row = None
    for i in range(len(cg_mat) - 1, -1, -1):
        if cg_mat[i][0] is not None:
            template_row = 3 + i
            break

    contadores      = {"novas": 0, "revisoes": 0}
    revisoes        = []
    escrita_cg      = {}   # linha → [steers, heifers, beef_cows, dairy_cows, bulls]
    escrita_cal     = {}   # linha → calves
    escrita_wt      = {}   # linha → [wt_cattle, wt_calves]
    linhas_novas    = set()  # linhas onde TODOS os dados eram None (novas de verdade)

    for (ano, mes), vals in sorted(dados.items()):
        linha = row_of(ano, mes)
        idx   = linha - 3
        eh_rev = (ano, mes) >= corte_anoMes
        rotulo = f"{mes:02d}/{ano}"

        # Valida alinhamento col B
        data_b = datas_col_b[idx] if 0 <= idx < len(datas_col_b) else None
        if isinstance(data_b, dt.datetime):
            if data_b.year != ano or data_b.month != mes:
                raise RuntimeError(
                    f"Divergência col B linha {linha}: "
                    f"esperado {mes:02d}/{ano}, encontrado {data_b.month:02d}/{data_b.year}"
                )

        # --- Contagens C:G ---
        atual_cg = cg_mat[idx] if idx < len(cg_mat) else [None] * 5
        novo_cg  = [vals["steers"], vals["heifers"], vals["beef_cows"], vals["dairy_cows"], vals["bulls"]]
        labels_cg = ["Steers", "Heifers", "Beef Cows", "Dairy Cows", "Bulls"]

        mudou_cg = False
        for label, atual, novo in zip(labels_cg, atual_cg, novo_cg):
            if novo is None:
                continue
            if atual is None:
                contadores["novas"] += 1
                mudou_cg = True
            elif eh_rev and _diff(atual, novo):
                contadores["revisoes"] += 1
                revisoes.append((rotulo, label, atual, novo))
                mudou_cg = True

        if mudou_cg:
            escrita_cg[linha] = novo_cg

        # --- Calves (I) ---
        atual_cal = cal_col[idx] if idx < len(cal_col) else None
        novo_cal  = vals["calves"]

        mudou_cal = False
        if novo_cal is not None:
            if atual_cal is None:
                contadores["novas"] += 1
                mudou_cal = True
            elif eh_rev and _diff(atual_cal, novo_cal):
                contadores["revisoes"] += 1
                revisoes.append((rotulo, "Calves", atual_cal, novo_cal))
                mudou_cal = True

        if mudou_cal:
            escrita_cal[linha] = novo_cal

        # --- Pesos P:Q ---
        atual_wt = wt_mat[idx] if idx < len(wt_mat) else [None, None]
        novo_wt  = [vals["wt_cattle"], vals["wt_calves"]]

        mudou_wt = False
        for label, atual, novo in zip(["Wt Cattle", "Wt Calves"], atual_wt, novo_wt):
            if novo is None:
                continue
            if atual is None:
                mudou_wt = True
            elif eh_rev and _diff(atual, novo):
                revisoes.append((rotulo, label, atual, novo))
                mudou_wt = True

        if mudou_wt:
            escrita_wt[linha] = novo_wt

        # Marca linha como nova se absolutamente todos os dados eram None
        row_era_vazio = (
            all(v is None for v in atual_cg)
            and atual_cal is None
            and all(v is None for v in atual_wt)
        )
        if row_era_vazio and (mudou_cg or mudou_cal or mudou_wt):
            linhas_novas.add(linha)

    # ---- Verificação de linhas em branco (gap fill) -------------------
    # Já coberta pelo loop acima: qualquer célula None + USDA tem valor → escreve.

    # ---- Escrita em blocos -------------------------------------------
    _escrever_blocos(sh, escrita_cg,  "C", "G", "contagens C:G")
    _escrever_col(sh, escrita_cal, "I", "calves I")
    _escrever_blocos(sh, escrita_wt,  "P", "Q", "pesos P:Q")

    # ---- Propagação de formatação e fórmulas para linhas novas -------
    if linhas_novas and template_row is not None:
        _propagar_formato_formulas(sh, linhas_novas, template_row)

    return {
        "salvar":       bool(escrita_cg or escrita_cal or escrita_wt),
        "contadores":   contadores,
        "revisoes":     revisoes,
        "linhas_novas": sorted(linhas_novas),
        "template_row": template_row,
        "meses_dados":  len(dados),
    }


# ============================================================
# ESCRITA EM BLOCOS
# ============================================================
def _blocos(linhas_sorted: list[int]) -> list[tuple[int, int]]:
    if not linhas_sorted:
        return []
    blocos = []
    ini = fim = linhas_sorted[0]
    for l in linhas_sorted[1:]:
        if l == fim + 1:
            fim = l
        else:
            blocos.append((ini, fim))
            ini = fim = l
    blocos.append((ini, fim))
    return blocos


def _escrever_blocos(sh, escrita: dict, col_ini: str, col_fim: str, label: str):
    if not escrita:
        return
    for ini, fim in _blocos(sorted(escrita)):
        matriz = [escrita[l] for l in range(ini, fim + 1) if l in escrita]
        sh.range(f"{col_ini}{ini}:{col_fim}{fim}").value = matriz
        print(f"  [{label}] linhas {ini}:{fim} ({fim - ini + 1} meses)")


def _escrever_col(sh, escrita: dict, col: str, label: str):
    if not escrita:
        return
    for ini, fim in _blocos(sorted(escrita)):
        valores = [[escrita[l]] for l in range(ini, fim + 1) if l in escrita]
        sh.range(f"{col}{ini}:{col}{fim}").value = valores
        print(f"  [{label}] linhas {ini}:{fim} ({fim - ini + 1} meses)")


# ============================================================
# PROPAGAÇÃO DE FORMATO E FÓRMULAS
# ============================================================
def _propagar_formato_formulas(sh, linhas_novas: set, template_row: int):
    """
    Para cada bloco de linhas novas:
      1. Copia formatação de toda a linha template.
      2. Propaga fórmula do Total (H) e % (J:N).
    """
    novas_ord = sorted(linhas_novas)
    for ini, fim in _blocos(novas_ord):
        # 1. Formatação (toda a linha, preserva número/cor/borda)
        sh.range(f"A{template_row}:Q{template_row}").api.Copy()
        sh.range(f"A{ini}:Q{fim}").api.PasteSpecial(Paste=-4122)   # xlPasteFormats

        # 2. Fórmula Total (col H)
        sh.range(f"{FORMULA_COL_TOTAL}{template_row}").api.Copy()
        sh.range(f"{FORMULA_COL_TOTAL}{ini}:{FORMULA_COL_TOTAL}{fim}").api.PasteSpecial(Paste=-4123)  # xlPasteFormulas

        # 3. Fórmulas % (J:N)
        sh.range(f"J{template_row}:N{template_row}").api.Copy()
        sh.range(f"J{ini}:N{fim}").api.PasteSpecial(Paste=-4123)   # xlPasteFormulas

        sh.book.app.api.CutCopyMode = False
        print(f"  [fmt+fórmulas] propagados para linhas {ini}:{fim}")


# ============================================================
# RELATÓRIO
# ============================================================
def print_relatorio(r: dict):
    print("\n== Atualização USDA — BeefUS Abates ==")
    print(f"Meses processados : {r['meses_dados']} ({ANO_INICIO}+)")
    print(f"Linhas novas      : {len(r['linhas_novas'])}")
    if r["linhas_novas"]:
        print(f"  (template usada : linha {r['template_row']})")

    c = r["contadores"]
    print(f"Células novas     : {c['novas']}")
    print(f"Revisões          : {c['revisoes']}")
    for mes, label, antigo, novo in r["revisoes"]:
        print(f"  - {mes} {label}: {antigo} → {novo}")

    print(f"\nStatus: {'SUCESSO' if r['salvar'] else 'SEM ALTERAÇÕES'}")


if __name__ == "__main__":
    main()
