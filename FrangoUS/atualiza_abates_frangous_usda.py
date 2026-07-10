"""
Atualização mensal ERS USDA — aba 'Production' de FrangoUS.xlsm

Fonte : ERS USDA – meat-statistics-tables-historical.xlsx
        https://www.ers.usda.gov/media/5529/meat-statistics-tables-historical.xlsx
        Sheet : SlaughterCounts-Full (índice 31 = Broilers, 1.000 cabeças)

Coluna escrita:
  H = USDA Commercial Slaughter (Broilers, em '000 heads)
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
XLSX_NAME  = "FrangoUS.xlsm"
SHEET_NAME = "Production"
ANO_INICIO = 2000   # linha 3 da aba = jan/2000

USDA_URL    = "https://www.ers.usda.gov/media/5529/meat-statistics-tables-historical.xlsx"
TIMEOUT     = 120
MAX_RETRIES = 3
RETRY_SLEEP = 5

REVISION_MONTHS = 12   # meses para verificar revisões retroativamente

# Índice (0-based) na sheet SlaughterCounts-Full do USDA
SC_BROILERS = 31

COL_BROILERS = "H"

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
    Retorna None para entradas trimestrais ('Jan-Mar 2026') ou inválidas.
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
    """Retorna { (ano, mes): valor_broilers_1000heads }"""
    wb = openpyxl.load_workbook(io.BytesIO(conteudo), read_only=True, data_only=True)
    ws = wb["SlaughterCounts-Full"]

    dados: dict[tuple, float] = {}
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i < 3:
            continue
        parsed = parse_usda_date(row[0])
        if parsed is None:
            continue
        ano, mes = parsed
        if ano < ANO_INICIO:
            continue
        val = row[SC_BROILERS]
        if val is not None:
            dados[(ano, mes)] = val

    print(f"  USDA: {len(dados)} meses lidos ({ANO_INICIO}+)")
    return dados


# ============================================================
# MAIN
# ============================================================
def main():
    path = localizar_planilha()

    print(f"[{dt.datetime.now():%Y-%m-%d %H:%M}] Iniciando atualização USDA — FrangoUS Commercial Slaughter...")

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

    # Lê col B (datas) e col H (valores atuais)
    datas_col_b  = sh.range(f"B3:B{last_row}").value
    broilers_col = sh.range(f"H3:H{last_row}").value

    if not isinstance(datas_col_b, list):
        datas_col_b = [datas_col_b]
    if not isinstance(broilers_col, list):
        broilers_col = [broilers_col]
    broilers_col = [v[0] if isinstance(v, (list, tuple)) else v for v in broilers_col]

    contadores = {"novas": 0, "revisoes": 0}
    revisoes   = []
    escrita    = {}   # linha → valor

    for (ano, mes), novo_val in sorted(dados.items()):
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

        atual = broilers_col[idx] if 0 <= idx < len(broilers_col) else None

        if atual is None:
            contadores["novas"] += 1
            escrita[linha] = novo_val
        elif eh_rev and _diff(atual, novo_val):
            contadores["revisoes"] += 1
            revisoes.append((rotulo, atual, novo_val))
            escrita[linha] = novo_val

    _escrever_col(sh, escrita, COL_BROILERS, "Commercial Slaughter H")
    _propagar_formato_col(sh, escrita, COL_BROILERS)

    return {
        "salvar":      bool(escrita),
        "contadores":  contadores,
        "revisoes":    revisoes,
        "meses_dados": len(dados),
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


def _escrever_col(sh, escrita: dict, col: str, label: str):
    if not escrita:
        return
    for ini, fim in _blocos(sorted(escrita)):
        valores = [[escrita[l]] for l in range(ini, fim + 1) if l in escrita]
        sh.range(f"{col}{ini}:{col}{fim}").value = valores
        print(f"  [{label}] linhas {ini}:{fim} ({fim - ini + 1} meses)")


def _propagar_formato_col(sh, escrita: dict, col: str):
    """Para cada bloco escrito, copia o formato da célula imediatamente acima."""
    if not escrita:
        return
    for ini, fim in _blocos(sorted(escrita)):
        origem = max(ini - 1, 3)   # nunca sobe acima da linha 3
        sh.range(f"{col}{origem}").api.Copy()
        sh.range(f"{col}{ini}:{col}{fim}").api.PasteSpecial(Paste=-4122)   # xlPasteFormats
        sh.book.app.api.CutCopyMode = False
        print(f"  [formato {col}] copiado de linha {origem} -> {ini}:{fim}")


# ============================================================
# RELATÓRIO
# ============================================================
def print_relatorio(r: dict):
    print("\n== Atualização USDA — FrangoUS Commercial Slaughter ==")
    print(f"Meses processados : {r['meses_dados']} ({ANO_INICIO}+)")
    c = r["contadores"]
    print(f"Células novas     : {c['novas']}")
    print(f"Revisões          : {c['revisoes']}")
    for mes, antigo, novo in r["revisoes"]:
        print(f"  - {mes}: {antigo} → {novo}")
    print(f"\nStatus: {'SUCESSO' if r['salvar'] else 'SEM ALTERAÇÕES'}")


if __name__ == "__main__":
    main()
