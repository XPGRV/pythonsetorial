"""
Atualização mensal PPI — aba 'Transformadores' de WEG - Setorial.xlsm

Fonte : FRED / St. Louis Fed
        Série PCU335311335311
        "PPI: Electric Power and Specialty Transformer Manufacturing"
Via   : xlwings (preserva links externos, formatação e demais abas intactas)

Coluna escrita na planilha:
  C = Index (valor mensal do PPI, começa em jan/2000 na linha 5)

Observações:
  - A coluna B já vem preenchida com as datas (mmm-yy) até 2030; usamos
    ela apenas para conferência de alinhamento.
  - Só gravamos células vazias ou que mudaram de valor (revisões do FRED).
"""

import io
import time
import datetime as dt
from pathlib import Path

import requests
import pandas as pd
import xlwings as xw


# ============================================================
# CONFIG
# ============================================================
DRIVE_ROOT = Path(
    r"G:\.shortcut-targets-by-id\11q1ngdqbySnDMCQSP4fiwPztT6ju2F4Z"
)
XLSX_NAME  = "WEG - Setorial.xlsm"
SHEET_NAME = "Transformadores"

SERIE      = "PCU335311335311"
FRED_URL   = f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={SERIE}"

ANO_INICIO = 2000          # jan/2000 = linha 5 (primeira linha de dados)
LINHA_JAN2000 = 5
COL_INDEX  = "C"

TIMEOUT     = 120
MAX_RETRIES = 3
RETRY_SLEEP = 5


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
    """Linha 5 = jan/2000."""
    return LINHA_JAN2000 + (ano - ANO_INICIO) * 12 + (mes - 1)


def _diff(a, b) -> bool:
    try:
        return abs(float(a) - float(b)) > 0.001
    except (TypeError, ValueError):
        return a != b


# ============================================================
# DOWNLOAD E PARSE DO CSV FRED
# ============================================================
def fetch_fred_csv() -> pd.DataFrame:
    """Baixa o CSV da série no FRED e retorna DataFrame bruto."""
    last_err = None
    for tentativa in range(1, MAX_RETRIES + 1):
        try:
            print(f"  Baixando CSV FRED (tentativa {tentativa}/{MAX_RETRIES})...")
            r = requests.get(FRED_URL, timeout=TIMEOUT)
            r.raise_for_status()
            conteudo = r.content
            break
        except Exception as e:
            last_err = e
            if tentativa < MAX_RETRIES:
                time.sleep(RETRY_SLEEP)
    else:
        raise RuntimeError(f"Download FRED falhou em {MAX_RETRIES} tentativas: {last_err}")

    df = pd.read_csv(io.BytesIO(conteudo))
    print(f"  CSV lido: {len(df):,} linhas")
    return df


def parse_fred(df: pd.DataFrame) -> dict:
    """
    Converte o CSV do FRED em {(ano, mes): valor}, a partir de ANO_INICIO.

    O FRED usa a 1ª coluna como data e a 2ª como valor da série; valores
    ausentes vêm como '.' (viram NaN e são ignorados).
    """
    col_data, col_valor = df.columns[0], df.columns[1]

    df = df.copy()
    df[col_data]  = pd.to_datetime(df[col_data], errors="coerce")
    df[col_valor] = pd.to_numeric(df[col_valor], errors="coerce")
    df = df[df[col_data].notna() & df[col_valor].notna()].copy()
    df = df[df[col_data].dt.year >= ANO_INICIO]

    resultado = {}
    for _, row in df.iterrows():
        d = row[col_data]
        resultado[(int(d.year), int(d.month))] = float(row[col_valor])

    return resultado


# ============================================================
# MAIN
# ============================================================
def main():
    path = localizar_planilha()

    print(f"[{dt.datetime.now():%Y-%m-%d %H:%M}] Iniciando atualização PPI Transformadores...")

    print("\n[1/3] Baixando e processando CSV do FRED...")
    df_bruto = fetch_fred_csv()
    dados    = parse_fred(df_bruto)
    print(f"  {len(dados)} meses com dados ({ANO_INICIO}+)")

    if not dados:
        print("Nenhum dado encontrado. Encerrando.")
        return

    print(f"\n[2/3] Abrindo {path.name} via xlwings...")
    app = xw.App(visible=False, add_book=False)
    try:
        app.api.AskToUpdateLinks = False
        app.display_alerts       = False
        app.screen_updating      = False

        wb = app.books.open(str(path), update_links=0, read_only=False)
        try:
            print("\n[3/3] Gravando dados na planilha...")
            resultado = processar(wb, dados)
            if resultado["salvar"]:
                wb.api.SaveLinkValues = True
                wb.save()
                print("Arquivo salvo.")
            else:
                print("Nada alterado — arquivo NÃO salvo.")
        finally:
            wb.close()
    finally:
        app.quit()

    print_relatorio(resultado)


# ============================================================
# PROCESSAMENTO
# ============================================================
def processar(wb, dados: dict) -> dict:
    sh = wb.sheets[SHEET_NAME]

    last_row    = sh.used_range.last_cell.row
    datas_col_b = sh.range(f"B{LINHA_JAN2000}:B{last_row}").value
    idx_vals    = sh.range(f"{COL_INDEX}{LINHA_JAN2000}:{COL_INDEX}{last_row}").value
    idx_lista   = idx_vals if isinstance(idx_vals, list) else [idx_vals]

    # Linha template = última linha JÁ preenchida na coluna C (antes desta
    # execução). Seu formato é propagado para as novas linhas escritas.
    template_row = None
    for i, v in enumerate(idx_lista):
        if v is not None:
            template_row = LINHA_JAN2000 + i
    print(f"  Linha template (formato): {template_row}")

    contadores       = {"escritas_novas": 0, "revisoes": 0}
    revisoes         = []
    linhas_novas     = []
    escrita_pendente: dict[int, float] = {}   # linha → valor

    for (ano, mes), novo in sorted(dados.items()):
        linha = row_of(ano, mes)
        idx   = linha - LINHA_JAN2000

        # Conferência de alinhamento com a coluna B (data real)
        data_b = datas_col_b[idx] if 0 <= idx < len(datas_col_b) else None
        if isinstance(data_b, dt.datetime):
            if data_b.year != ano or data_b.month != mes:
                raise RuntimeError(
                    f"Divergência col B linha {linha}: "
                    f"esperado {mes:02d}/{ano}, "
                    f"encontrado {data_b.month:02d}/{data_b.year}"
                )
        elif data_b is None:
            # Fora do range pré-preenchido da planilha — ignora
            continue

        atual = idx_lista[idx] if 0 <= idx < len(idx_lista) else None
        rotulo = f"{mes:02d}/{ano}"

        if atual is None:
            contadores["escritas_novas"] += 1
            linhas_novas.append(linha)
            escrita_pendente[linha] = novo
        elif _diff(atual, novo):
            contadores["revisoes"] += 1
            revisoes.append((rotulo, atual, novo))
            escrita_pendente[linha] = novo

    # ---- escrita em bloco (linhas contíguas = 1 chamada COM) --------
    if escrita_pendente:
        linhas_ord  = sorted(escrita_pendente)
        blocos: list[list[int]] = []
        bloco_atual = [linhas_ord[0]]
        for l in linhas_ord[1:]:
            if l == bloco_atual[-1] + 1:
                bloco_atual.append(l)
            else:
                blocos.append(bloco_atual)
                bloco_atual = [l]
        blocos.append(bloco_atual)

        for bloco in blocos:
            coluna  = [[escrita_pendente[l]] for l in bloco]   # vetor coluna
            ini, fim = bloco[0], bloco[-1]
            sh.range(f"{COL_INDEX}{ini}:{COL_INDEX}{fim}").value = coluna
            print(f"  Escritas linhas {ini}:{fim} ({len(bloco)} meses)")

    # ---- propagar formato da linha template p/ as linhas novas -------
    if linhas_novas and template_row is not None:
        novas_ord = sorted(linhas_novas)

        # Blocos contíguos (normalmente um só)
        blocos_fmt: list[tuple[int, int]] = []
        ini = fim = novas_ord[0]
        for l in novas_ord[1:]:
            if l == fim + 1:
                fim = l
            else:
                blocos_fmt.append((ini, fim))
                ini = fim = l
        blocos_fmt.append((ini, fim))

        for ini, fim in blocos_fmt:
            sh.range(f"{COL_INDEX}{template_row}").api.Copy()
            sh.range(f"{COL_INDEX}{ini}:{COL_INDEX}{fim}").api.PasteSpecial(
                Paste=-4122  # xlPasteFormats
            )
            print(f"  Formato da linha {template_row} aplicado em {ini}:{fim}")

        sh.book.app.api.CutCopyMode = False

    return {
        "salvar":       sum(contadores.values()) > 0,
        "contadores":   contadores,
        "revisoes":     revisoes,
        "linhas_novas": linhas_novas,
        "template_row": template_row,
        "meses_dados":  len(dados),
    }


# ============================================================
# RELATÓRIO
# ============================================================
def print_relatorio(r: dict):
    print("\n== Atualização PPI — Transformadores (WEG) ==")
    print(f"Meses processados: {r['meses_dados']} ({ANO_INICIO}+)")

    c = r["contadores"]
    print(f"\nEscritas novas : {c['escritas_novas']}")
    print(f"Revisões       : {c['revisoes']}")
    for mes, antigo, novo in r["revisoes"]:
        print(f"  - {mes}: {antigo} → {novo}")

    if r["linhas_novas"]:
        print(f"\nPrimeira linha nova: {r['linhas_novas'][0]} | "
              f"última: {r['linhas_novas'][-1]}")
    print(f"Linha template (formato): {r['template_row']}")

    print(f"\nStatus: {'SUCESSO' if r['salvar'] else 'SEM ALTERAÇÕES'}")


if __name__ == "__main__":
    main()
