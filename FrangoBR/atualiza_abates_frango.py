"""
Atualização mensal SIF — aba 'Production' de FrangoBR.xlsm

Fonte : MAPA/SIGSIF – sigsifrelatorioabates.csv
Via   : xlwings (preserva links externos, formatação e demais abas intactas)

Colunas SIF escritas na planilha:
  F = Frango       (total de cabeças: Frango + Frango Especial, machos + fêmeas)
  G = Galo+Galinha (total de cabeças: Galo + Galinha, machos + fêmeas)
  H = Total        (fórmula =F+G — propagada do template, não escrita aqui)
"""

import io
import re
import sys
import time
import datetime as dt
from copy import copy
from pathlib import Path

import requests
import pandas as pd
import openpyxl
from openpyxl.utils import column_index_from_string

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import drive_io


_REF = re.compile(r"(?<![A-Za-z0-9_])(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z_(])")
def shift_formula(formula, delta):
    if not isinstance(formula, str) or not formula.startswith("="):
        return formula
    def repl(m):
        cabs, col, rabs, row = m.groups()
        return m.group(0) if rabs == "$" else f"{cabs}{col}{rabs}{int(row)+delta}"
    return _REF.sub(repl, formula)


# ============================================================
# CONFIG  —  ajustar antes de rodar
# ============================================================
DRIVE_ROOT = Path(
    r"G:\.shortcut-targets-by-id\11q1ngdqbySnDMCQSP4fiwPztT6ju2F4Z"
)
XLSX_NAME  = "FrangoBR.xlsm"
SHEET_NAME = "Production"
ANO_INICIO = 2005

SIF_URL = (
    "https://dados.agricultura.gov.br/dataset/062166e3-b515-4274-8e7d-68aadd64b820"
    "/resource/341dc717-4716-42ab-b189-c8d7a9d2a1ba/download/sigsifrelatorioabates.csv"
)

TIMEOUT     = 120
MAX_RETRIES = 3
RETRY_SLEEP = 5

# Colunas na planilha
COL_FRANGO      = "F"   # Frango + Frango Especial
COL_GALINHA_GAL = "G"   # Galinha + Galo
COL_TOTAL       = "H"   # Fórmula =F+G (propagada do template, não escrita pelo script)

# Coluna com datas mensais (col A = trimestres, col B = meses)
COL_DATAS = "B"

# Linha base: linha 3 = jan/2000
ANO_BASE = 2000
ROW_BASE = 3

# Categorias por coluna (QTD_MACHO + QTD_FEMEA somados para cada grupo)
CATS_FRANGO      = {"Frango", "Frango Especial"}
CATS_GALINHA_GAL = {"Galinha", "Galo"}


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
    return ROW_BASE + (ano - ANO_BASE) * 12 + (mes - 1)


def _diff(a, b) -> bool:
    try:
        return abs(float(a) - float(b)) > 0.5
    except (TypeError, ValueError):
        return a != b


# ============================================================
# DOWNLOAD E PARSE DO CSV SIF
# ============================================================
def fetch_sif_csv() -> pd.DataFrame:
    last_err = None
    for tentativa in range(1, MAX_RETRIES + 1):
        try:
            print(f"  Baixando CSV SIF (tentativa {tentativa}/{MAX_RETRIES})...")
            r = requests.get(SIF_URL, timeout=TIMEOUT)
            r.raise_for_status()
            conteudo = r.content
            last_modified = r.headers.get("Last-Modified", "")
            if last_modified:
                from email.utils import parsedate_to_datetime
                lm_dt = parsedate_to_datetime(last_modified).replace(tzinfo=None)
                idade = (dt.datetime.now(dt.timezone.utc).replace(tzinfo=None) - lm_dt).days
                aviso = " ⚠️  MAPA ainda não atualizou este mês!" if idade > 35 else ""
                print(f"  CSV gerado pelo MAPA em: {lm_dt:%d/%m/%Y %H:%M} UTC  ({idade} dias atrás){aviso}")
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
    Filtra às categorias de aves selecionadas e agrega por (ano, mes) nacional.

    Retorna: {(ano, mes): {"frango": int, "galinha_galo": int}}
    """
    df = df.copy()
    df.columns      = df.columns.str.strip()
    df["CATEGORIA"] = df["CATEGORIA"].str.strip()
    df["MES_ANO"]   = df["MES_ANO"].str.strip()

    datas      = pd.to_datetime(df["MES_ANO"], format="%m/%Y", errors="coerce")
    df["mes"]  = datas.dt.month
    df["ano"]  = datas.dt.year
    df = df[df["mes"].notna() & df["ano"].notna()].copy()
    df["mes"]  = df["mes"].astype(int)
    df["ano"]  = df["ano"].astype(int)

    todas_cats = CATS_FRANGO | CATS_GALINHA_GAL
    df = df[(df["ano"] >= ANO_INICIO) & df["CATEGORIA"].isin(todas_cats)].copy()

    df["QTD_MACHO"] = pd.to_numeric(df["QTD_MACHO"], errors="coerce").fillna(0).astype(int)
    df["QTD_FEMEA"] = pd.to_numeric(df["QTD_FEMEA"], errors="coerce").fillna(0).astype(int)
    df["TOTAL"]     = df["QTD_MACHO"] + df["QTD_FEMEA"]

    df["GRUPO"] = df["CATEGORIA"].apply(
        lambda c: "frango" if c in CATS_FRANGO else "galinha_galo"
    )

    agrupado = (
        df.groupby(["ano", "mes", "GRUPO"])["TOTAL"]
        .sum()
        .unstack(fill_value=0)
        .reset_index()
    )

    for col in ("frango", "galinha_galo"):
        if col not in agrupado.columns:
            agrupado[col] = 0

    resultado = {}
    for _, row in agrupado.iterrows():
        resultado[(int(row["ano"]), int(row["mes"]))] = {
            "frango":      int(row["frango"]),
            "galinha_galo": int(row["galinha_galo"]),
        }

    return resultado


# ============================================================
# MAIN
# ============================================================
def main():
    fid = drive_io.find_file_id(XLSX_NAME)
    path = drive_io.download_to_temp(fid)

    print(f"[{dt.datetime.now():%Y-%m-%d %H:%M}] Iniciando atualização SIF Frango...")

    print("\n[1/3] Baixando e processando CSV do SIF...")
    df_bruto = fetch_sif_csv()
    dados    = agregar_sif(df_bruto)
    print(f"  {len(dados)} meses com dados de aves ({ANO_INICIO}+)")

    if not dados:
        print("Nenhum dado encontrado. Encerrando.")
        return

    print("\n[2/3] Abrindo FrangoBR.xlsm via openpyxl...")
    wb = openpyxl.load_workbook(path, keep_vba=True, data_only=False)
    try:
        print("\n[3/3] Gravando dados na planilha...")
        resultado = processar(wb, dados)
        if resultado["salvar"]:
            wb.save(path)
            drive_io.update_file(fid, path)   # sobe de volta pro Drive (in-place)
            print("Arquivo salvo.")
        else:
            print("Nada alterado — arquivo NAO salvo.")
    finally:
        wb.close()

    print_relatorio(resultado)


# ============================================================
# PROCESSAMENTO
# ============================================================
def processar(wb, dados: dict) -> dict:
    ws = wb[SHEET_NAME]
    cB = column_index_from_string(COL_DATAS)
    cF = column_index_from_string(COL_FRANGO)
    cG = column_index_from_string(COL_GALINHA_GAL)
    cH = column_index_from_string(COL_TOTAL)

    last_row  = ws.max_row
    datas_col = [ws.cell(r, cB).value for r in range(ROW_BASE, last_row + 1)]

    # Estado atual de F:G (cada linha = [frango, galinha_galo])
    sif_lista = [[ws.cell(r, cF).value, ws.cell(r, cG).value]
                 for r in range(ROW_BASE, last_row + 1)]

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

    # Linha template = última linha com dados em colunas de referência SIDRA (ex: C:G)
    heads_vals = [[ws.cell(r, c).value for c in range(3, 8)]
                  for r in range(ROW_BASE, last_row + 1)]
    _, u_h, _    = diagnostico(heads_vals)
    template_row = (ROW_BASE + u_h) if u_h is not None else None
    print(f"  Linha template: {template_row}")

    contadores       = {"escritas_novas": 0, "revisoes": 0}
    revisoes         = []
    linhas_novas     = set()
    escrita_pendente: dict[int, list] = {}   # linha → [frango, galinha_galo]

    for (ano, mes), vals in sorted(dados.items()):
        linha  = row_of(ano, mes)
        idx_b  = linha - ROW_BASE
        data_b = datas_col[idx_b] if 0 <= idx_b < len(datas_col) else None

        if isinstance(data_b, dt.datetime):
            if data_b.year != ano or data_b.month != mes:
                raise RuntimeError(
                    f"Divergência col {COL_DATAS} linha {linha}: "
                    f"esperado {mes:02d}/{ano}, "
                    f"encontrado {data_b.month:02d}/{data_b.year}"
                )

        rotulo  = f"{mes:02d}/{ano}"
        sif_idx = linha - ROW_BASE

        if sif_idx < len(sif_lista):
            atual_linha = sif_lista[sif_idx]
            if not isinstance(atual_linha, (list, tuple)):
                atual_linha = [atual_linha, None]
        else:
            atual_linha = [None, None]

        tinha = any(v is not None for v in atual_linha)
        if not tinha:
            linhas_novas.add(linha)

        novo_frango = vals["frango"]
        novo_galgal = vals["galinha_galo"]
        atual_f, atual_g = atual_linha[0], atual_linha[1]

        linha_mudou = False
        for label, atual, novo in [
            ("Frango",       atual_f, novo_frango),
            ("Galo+Galinha", atual_g, novo_galgal),
        ]:
            if atual is None:
                contadores["escritas_novas"] += 1
                linha_mudou = True
            elif _diff(atual, novo):
                contadores["revisoes"] += 1
                revisoes.append((rotulo, label, atual, novo))
                linha_mudou = True

        if linha_mudou:
            escrita_pendente[linha] = [novo_frango, novo_galgal]

    # Escrita em bloco (linhas contíguas = 1 chamada COM)
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
            ini, fim = bloco[0], bloco[-1]
            for l in bloco:
                f_val, g_val = escrita_pendente[l]
                ws.cell(l, cF).value = f_val
                ws.cell(l, cG).value = g_val
            print(f"  Escritas linhas {ini}:{fim} ({len(bloco)} meses)")

    # Propagar formato (V:W) e fórmula (X) para linhas novas
    if linhas_novas and template_row is not None:
        novas_ord = sorted(linhas_novas)

        blocos_fmt: list[tuple[int, int]] = []
        ini = fim = novas_ord[0]
        for l in novas_ord[1:]:
            if l == fim + 1:
                fim = l
            else:
                blocos_fmt.append((ini, fim))
                ini = fim = l
        blocos_fmt.append((ini, fim))

        h_tmpl = ws.cell(template_row, cH).value   # formula Total (=IFERROR(G+F,...))
        for ini, fim in blocos_fmt:
            for l in range(ini, fim + 1):
                # Formato de F, G e H copiado do template
                for c in (cF, cG, cH):
                    src = ws.cell(template_row, c)
                    dst = ws.cell(l, c)
                    dst.number_format = src.number_format
                    dst.font = copy(src.font)
                    dst.fill = copy(src.fill)
                    dst.border = copy(src.border)
                    dst.alignment = copy(src.alignment)
                # Formula H (Total = G+F), com linhas deslocadas
                ws.cell(l, cH).value = shift_formula(h_tmpl, l - template_row)

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
    print("\n== Atualização SIF — Abates de Aves ==")
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
