"""
Atualização mensal ERS USDA — aba 'Production' de FrangoUS.xlsm

Fonte : ERS USDA – meat-statistics-tables-historical.xlsx
        https://www.ers.usda.gov/media/5529/meat-statistics-tables-historical.xlsx
        Sheet : SlaughterCounts-Full (índice 31 = Broilers, 1.000 cabeças)

Coluna escrita:
  H = USDA Commercial Slaughter (Broilers, em '000 heads)
"""

import io
import sys
import time
import datetime as dt
from copy import copy
from pathlib import Path

import requests
import openpyxl
from openpyxl.utils import column_index_from_string

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import usda_ers


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
# Nomes do cabecalho, nao posicoes: em out/2026 o USDA removeu uma coluna
# de cada bloco e as posicoes fixas passaram a apontar para o vizinho.
SC_BROILERS_NOME = "Broilers"

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
    wb = usda_ers.abrir(conteudo)
    ws = usda_ers.aba(wb, "SlaughterCounts", "SlaughterCounts-Full")
    col = usda_ers.colunas(ws, SC_BROILERS_NOME)[SC_BROILERS_NOME]

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
        val = row[col]
        if val is not None:
            dados[(ano, mes)] = val

    print(f"  USDA: {len(dados)} meses lidos ({ANO_INICIO}+)")
    return dados


# ============================================================
# MAIN
# ============================================================
def main():
    # Baixa a planilha do Google Drive (em vez do drive montado em G:\).
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    import drive_io
    fid  = drive_io.find_file_id(XLSX_NAME)
    path = drive_io.download_to_temp(fid)

    print(f"[{dt.datetime.now():%Y-%m-%d %H:%M}] Iniciando atualização USDA — FrangoUS Commercial Slaughter...")

    print("\n[1/3] Baixando e processando Excel do USDA...")
    conteudo = download_usda()
    dados    = parse_usda(conteudo)

    if not dados:
        print("Nenhum dado encontrado. Encerrando.")
        return

    print("\n[2/3] Abrindo FrangoUS.xlsm via openpyxl...")
    wb = openpyxl.load_workbook(path, keep_vba=True, data_only=False)
    try:
        print("\n[3/3] Gravando dados na planilha...")
        resultado = processar(wb, dados)
        if resultado["salvar"]:
            wb.save(path)
            drive_io.update_file(fid, path)   # sobe de volta pro Drive (in-place)
            print("Arquivo salvo.")
        else:
            print("Nada alterado — arquivo NÃO salvo.")
    finally:
        wb.close()

    print_relatorio(resultado)


# ============================================================
# PROCESSAMENTO
# ============================================================
def processar(wb, dados: dict) -> dict:
    ws = wb[SHEET_NAME]

    hoje         = dt.date.today()
    corte_rev    = hoje - dt.timedelta(days=REVISION_MONTHS * 31)
    corte_anoMes = (corte_rev.year, corte_rev.month)

    last_row = ws.max_row

    # Lê col B (datas) e col H (valores atuais), a partir da linha 3.
    datas_col_b  = [ws.cell(row=r, column=2).value for r in range(3, last_row + 1)]
    broilers_col = [ws.cell(row=r, column=8).value for r in range(3, last_row + 1)]

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

    _escrever_col(ws, escrita, COL_BROILERS, "Commercial Slaughter H")
    _propagar_formato_col(ws, escrita, COL_BROILERS)

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


def _escrever_col(ws, escrita: dict, col: str, label: str):
    if not escrita:
        return
    ci = column_index_from_string(col)
    for ini, fim in _blocos(sorted(escrita)):
        for l in range(ini, fim + 1):
            if l in escrita:
                ws.cell(row=l, column=ci).value = escrita[l]
        print(f"  [{label}] linhas {ini}:{fim} ({fim - ini + 1} meses)")


def _propagar_formato_col(ws, escrita: dict, col: str):
    """Para cada bloco escrito, copia o formato da célula imediatamente acima."""
    if not escrita:
        return
    ci = column_index_from_string(col)
    for ini, fim in _blocos(sorted(escrita)):
        origem = max(ini - 1, 3)   # nunca sobe acima da linha 3
        src = ws.cell(row=origem, column=ci)
        for l in range(ini, fim + 1):
            dst = ws.cell(row=l, column=ci)
            dst.number_format = src.number_format
            dst.font          = copy(src.font)
            dst.fill          = copy(src.fill)
            dst.border        = copy(src.border)
            dst.alignment     = copy(src.alignment)
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
