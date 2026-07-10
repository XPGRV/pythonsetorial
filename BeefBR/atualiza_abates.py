"""
Atualização mensal SIF — aba 'Abates' de BeefBR.xlsm

Fonte : MAPA/SIGSIF – sigsifrelatorioabates.csv
Via   : xlwings (preserva links externos, formatação e demais abas intactas)

Colunas SIF escritas na planilha:
  V = Machos (QTD_MACHO agregado de todas as categorias bovinas)
  W = Fêmeas (QTD_FEMEA agregado de todas as categorias bovinas)

Categorias bovinas consideradas:
  Bovino, Touro/Touruno, Vaca,
  Novilho Precoce, Novilhao, Novilho Intermediario,
  Novilha Precoce, Novilha Intermediaria, Novilhona,
  Vitelo
  (Búfalos, Nonato e demais são ignorados)
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
XLSX_NAME  = "BeefBR.xlsm"
SHEET_NAME = "Abates"
ANO_INICIO = 2021

SIF_URL = (
    "https://dados.agricultura.gov.br/dataset/062166e3-b515-4274-8e7d-68aadd64b820"
    "/resource/341dc717-4716-42ab-b189-c8d7a9d2a1ba/download/sigsifrelatorioabates.csv"
)

TIMEOUT     = 120
MAX_RETRIES = 3
RETRY_SLEEP = 5

# Colunas SIF na planilha
COL_MACHOS = "V"
COL_FEMEAS = "W"

# Categorias bovinas que entram na soma (todas as demais são ignoradas)
CATS_BOVINAS = {
    "Bovino",
    "Touro/Touruno",
    "Vaca",
    "Novilho Precoce",
    "Novilhao",
    "Novilho Intermediario",
    "Novilha Precoce",
    "Novilha Intermediaria",
    "Novilhona",
    "Vitelo",
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
    """Linha 3 = jan/97."""
    return 3 + (ano - 1997) * 12 + (mes - 1)


def _diff(a, b) -> bool:
    try:
        return abs(float(a) - float(b)) > 0.5
    except (TypeError, ValueError):
        return a != b


# ============================================================
# DOWNLOAD E PARSE DO CSV SIF
# ============================================================
def fetch_sif_csv() -> pd.DataFrame:
    """Baixa o CSV do SIF e retorna DataFrame bruto."""
    last_err = None
    for tentativa in range(1, MAX_RETRIES + 1):
        try:
            print(f"  Baixando CSV SIF (tentativa {tentativa}/{MAX_RETRIES})...")
            r = requests.get(SIF_URL, timeout=TIMEOUT)
            r.raise_for_status()
            conteudo = r.content
            break
        except Exception as e:
            last_err = e
            if tentativa < MAX_RETRIES:
                time.sleep(RETRY_SLEEP)
    else:
        raise RuntimeError(f"Download SIF falhou em {MAX_RETRIES} tentativas: {last_err}")

    for enc in ("utf-8", "latin-1", "cp1252"):
        try:
            df = pd.read_csv(io.BytesIO(conteudo), sep=";", encoding=enc, dtype=str)
            print(f"  CSV lido ({enc}): {len(df):,} linhas")
            return df
        except Exception:
            continue

    raise RuntimeError("Não foi possível decodificar o CSV do SIF.")


def agregar_sif(df: pd.DataFrame) -> dict:
    """
    Filtra às categorias bovinas e agrega por (ano, mes) nacional.

    Retorna: {(ano, mes): {"machos": int, "femeas": int}}
    """
    df = df.copy()
    df.columns      = df.columns.str.strip()
    df["CATEGORIA"] = df["CATEGORIA"].str.strip()
    df["MES_ANO"]   = df["MES_ANO"].str.strip()

    # Parse de data — suporta M/YYYY e MM/YYYY
    datas      = pd.to_datetime(df["MES_ANO"], format="%m/%Y", errors="coerce")
    df["mes"]  = datas.dt.month
    df["ano"]  = datas.dt.year
    df = df[df["mes"].notna() & df["ano"].notna()].copy()
    df["mes"]  = df["mes"].astype(int)
    df["ano"]  = df["ano"].astype(int)

    # Filtra período e categorias
    df = df[(df["ano"] >= ANO_INICIO) & df["CATEGORIA"].isin(CATS_BOVINAS)].copy()

    df["QTD_MACHO"] = pd.to_numeric(df["QTD_MACHO"], errors="coerce").fillna(0).astype(int)
    df["QTD_FEMEA"] = pd.to_numeric(df["QTD_FEMEA"], errors="coerce").fillna(0).astype(int)

    agrupado = (
        df.groupby(["ano", "mes"])
        .agg(machos=("QTD_MACHO", "sum"), femeas=("QTD_FEMEA", "sum"))
        .reset_index()
    )

    resultado = {}
    for _, row in agrupado.iterrows():
        resultado[(int(row["ano"]), int(row["mes"]))] = {
            "machos": int(row["machos"]),
            "femeas": int(row["femeas"]),
        }

    return resultado


# ============================================================
# MAIN
# ============================================================
def main():
    path = localizar_planilha()

    print(f"[{dt.datetime.now():%Y-%m-%d %H:%M}] Iniciando atualização SIF...")

    print("\n[1/3] Baixando e processando CSV do SIF...")
    df_bruto = fetch_sif_csv()
    dados    = agregar_sif(df_bruto)
    print(f"  {len(dados)} meses com dados bovinos ({ANO_INICIO}+)")

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
    datas_col_b = sh.range(f"B3:B{last_row}").value

    # Lê estado atual das duas colunas SIF (V:W) de uma vez
    sif_vals  = sh.range(f"{COL_MACHOS}3:{COL_FEMEAS}{last_row}").value
    sif_lista = sif_vals if isinstance(sif_vals, list) else []

    def diagnostico(matriz):
        primeiro = ultimo = None
        preenchidas = set()
        for i, linha in enumerate(matriz or []):
            if not isinstance(linha, (list, tuple)):
                linha = [linha]
            if any(v is not None for v in linha):
                if primeiro is None:
                    primeiro = i
                ultimo = i
                preenchidas.add(i)
        gap = (
            primeiro is None
            or primeiro != 0
            or (ultimo is not None and len(preenchidas) != ultimo - primeiro + 1)
        )
        return primeiro, ultimo, gap

    p_sif, u_sif, gap_sif = diagnostico(sif_lista)
    print(f"  SIF (V:W): primeiro={p_sif}, último={u_sif}, gap={gap_sif}")

    # Linha template = última linha com dados em heads SIDRA (C:G)
    heads_vals   = sh.range(f"C3:G{last_row}").value
    _, u_h, _    = diagnostico(heads_vals)
    template_row = (3 + u_h) if u_h is not None else None
    print(f"  Linha template: {template_row}")

    # ---- compara valores e monta escrita pendente -------------------
    contadores       = {"escritas_novas": 0, "revisoes": 0}
    revisoes         = []
    linhas_novas     = set()
    escrita_pendente: dict[int, list] = {}   # linha → [machos, femeas]

    for (ano, mes), vals in sorted(dados.items()):
        linha   = row_of(ano, mes)
        idx_b   = linha - 3
        data_b  = datas_col_b[idx_b] if 0 <= idx_b < len(datas_col_b) else None

        if isinstance(data_b, dt.datetime):
            if data_b.year != ano or data_b.month != mes:
                raise RuntimeError(
                    f"Divergência col B linha {linha}: "
                    f"esperado {mes:02d}/{ano}, "
                    f"encontrado {data_b.month:02d}/{data_b.year}"
                )

        rotulo  = f"{mes:02d}/{ano}"
        sif_idx = linha - 3

        # Valores atuais (V, W) para essa linha
        if sif_idx < len(sif_lista):
            atual_linha = sif_lista[sif_idx]
            if not isinstance(atual_linha, (list, tuple)):
                atual_linha = [atual_linha, None]
        else:
            atual_linha = [None, None]

        tinha = any(v is not None for v in atual_linha)
        if not tinha:
            linhas_novas.add(linha)

        novo_m = vals["machos"]
        novo_f = vals["femeas"]
        atual_m, atual_f = atual_linha[0], atual_linha[1]

        linha_mudou = False
        for label, atual, novo in [("Machos", atual_m, novo_m), ("Fêmeas", atual_f, novo_f)]:
            if atual is None:
                contadores["escritas_novas"] += 1
                linha_mudou = True
            elif _diff(atual, novo):
                contadores["revisoes"] += 1
                revisoes.append((rotulo, label, atual, novo))
                linha_mudou = True

        if linha_mudou:
            escrita_pendente[linha] = [novo_m, novo_f]

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
            matriz  = [escrita_pendente[l] for l in bloco]
            ini, fim = bloco[0], bloco[-1]
            sh.range(f"{COL_MACHOS}{ini}:{COL_FEMEAS}{fim}").value = matriz
            print(f"  Escritas linhas {ini}:{fim} ({len(bloco)} meses)")

    # ---- propagar formato e fórmulas para linhas novas ---------------
    if linhas_novas and template_row is not None:
        novas_ord = sorted(linhas_novas)

        # Detecta blocos contíguos (normalmente um único bloco)
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
            # Formato de V:W
            sh.range(f"{COL_MACHOS}{template_row}:{COL_FEMEAS}{template_row}").api.Copy()
            sh.range(f"{COL_MACHOS}{ini}:{COL_FEMEAS}{fim}").api.PasteSpecial(Paste=-4122)  # xlPasteFormats

            # Fórmulas de X:Z (Excel ajusta referências automaticamente)
            sh.range(f"X{template_row}:Z{template_row}").api.Copy()
            sh.range(f"X{ini}:Z{fim}").api.PasteSpecial(Paste=-4123)  # xlPasteFormulas

        sh.book.app.api.CutCopyMode = False

    return {
        "salvar":       sum(contadores.values()) > 0,
        "contadores":   contadores,
        "revisoes":     revisoes,
        "linhas_novas": sorted(linhas_novas),
        "template_row": template_row,
        "meses_dados":  len(dados),
    }


# ============================================================
# RELATÓRIO
# ============================================================
def print_relatorio(r: dict):
    print("\n== Atualização SIF — Abates Bovinos ==")
    print(f"Meses processados: {r['meses_dados']} ({ANO_INICIO}+)")
    print(f"\nLinhas novas (V:W vazios): {len(r['linhas_novas'])}")
    for linha in r["linhas_novas"]:
        print(f"  - linha {linha}")
    print(f"Linha template usada: {r['template_row']}")

    c = r["contadores"]
    print(f"\nEscritas novas : {c['escritas_novas']}")
    print(f"Revisões       : {c['revisoes']}")
    for mes, label, antigo, novo in r["revisoes"]:
        print(f"  - {mes} {label}: {antigo:,.0f} → {novo:,.0f}")

    print(f"\nStatus: {'SUCESSO' if r['salvar'] else 'SEM ALTERAÇÕES'}")


if __name__ == "__main__":
    main()
