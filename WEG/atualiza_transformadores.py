"""
Atualização mensal PPI — coluna C da aba 'Transformadores' (WEG - Setorial.xlsm)

Fonte : FRED / St. Louis Fed — série PCU335311335311
        "PPI: Electric Power and Specialty Transformer Manufacturing"
Via   : xlwings (preserva macros, links externos e demais abas intactas)

Escopo: escreve APENAS a coluna C (Index), a partir de jan/2000 na linha 5.
        As demais colunas da pasta (SECEX, Solar, Peers) vêm de outras fontes
        e não são tocadas por este script.

Comportamento:
  - só grava célula vazia ou cujo valor mudou (revisão do FRED);
  - novas linhas herdam a formatação da última linha já preenchida.
"""

import io
import time
import datetime as dt
from pathlib import Path

import sys
from copy import copy

import requests
import pandas as pd
import openpyxl
from openpyxl.utils import column_index_from_string

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import drive_io


# ============================================================
# CONFIG
# ============================================================
DRIVE_ROOT = Path(
    r"G:\.shortcut-targets-by-id\11q1ngdqbySnDMCQSP4fiwPztT6ju2F4Z"
)
XLSX_NAME  = "WEG - Setorial.xlsm"
SHEET_NAME = "Transformadores"

SERIE    = "PCU335311335311"
FRED_URL = f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={SERIE}"

ANO_INICIO    = 2000   # jan/2000 ...
LINHA_JAN2000 = 5      # ... na linha 5
COL           = "C"    # coluna Index

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
            f"Foram encontradas várias planilhas chamadas {XLSX_NAME!r}:\n{caminhos}"
        )

    return encontrados[0]


def row_of(ano: int, mes: int) -> int:
    """Linha 5 = jan/2000."""
    return LINHA_JAN2000 + (ano - ANO_INICIO) * 12 + (mes - 1)


def _mudou(atual, novo) -> bool:
    try:
        return abs(float(atual) - float(novo)) > 0.001
    except (TypeError, ValueError):
        return True


def _coluna(valores, n: int) -> list:
    """Normaliza o retorno do xlwings para uma lista de n elementos."""
    if not isinstance(valores, list):
        valores = [valores]
    return valores + [None] * (n - len(valores))


def blocos_contiguos(linhas) -> list[tuple[int, int]]:
    """Agrupa linhas em blocos contíguos: [12, 13, 14, 20] -> [(12,14), (20,20)]."""
    blocos: list[list[int]] = []
    for linha in sorted(linhas):
        if blocos and linha == blocos[-1][1] + 1:
            blocos[-1][1] = linha
        else:
            blocos.append([linha, linha])
    return [(ini, fim) for ini, fim in blocos]


# ============================================================
# FRED
# ============================================================
def baixar_fred() -> dict:
    """Baixa a série no FRED e retorna {(ano, mes): valor} a partir de ANO_INICIO."""
    erro = None
    for tentativa in range(1, MAX_RETRIES + 1):
        try:
            print(f"  Baixando CSV FRED (tentativa {tentativa}/{MAX_RETRIES})...")
            r = requests.get(FRED_URL, timeout=TIMEOUT)
            r.raise_for_status()
            break
        except Exception as e:
            erro = e
            if tentativa < MAX_RETRIES:
                time.sleep(RETRY_SLEEP)
    else:
        raise RuntimeError(f"Download FRED falhou em {MAX_RETRIES} tentativas: {erro}")

    df = pd.read_csv(io.BytesIO(r.content))
    col_data, col_valor = df.columns[0], df.columns[1]

    # Valores ausentes vêm como '.' no FRED — viram NaN e são descartados
    df[col_data]  = pd.to_datetime(df[col_data], errors="coerce")
    df[col_valor] = pd.to_numeric(df[col_valor], errors="coerce")
    df = df[df[col_data].notna() & df[col_valor].notna()]
    df = df[df[col_data].dt.year >= ANO_INICIO]

    return {
        (d.year, d.month): float(v)
        for d, v in zip(df[col_data], df[col_valor])
    }


# ============================================================
# PROCESSAMENTO
# ============================================================
def processar(sh, dados: dict) -> dict:
    # Fim do grid de datas da coluna B (a aba tem max_row inflado, por isso
    # varremos ate a primeira lacuna) e recorte no ultimo mes disponivel no FRED.
    col_idx = column_index_from_string(COL)
    fim_grid = LINHA_JAN2000
    for r in range(LINHA_JAN2000, sh.max_row + 1):
        if sh.cell(r, 2).value is None:
            break
        fim_grid = r
    fim = min(row_of(*max(dados)), fim_grid)
    n   = fim - LINHA_JAN2000 + 1

    datas  = [sh.cell(r, 2).value for r in range(LINHA_JAN2000, fim + 1)]
    atuais = [sh.cell(r, col_idx).value for r in range(LINHA_JAN2000, fim + 1)]

    # Template de formatação = última linha já preenchida na coluna C
    preenchidas  = [i for i, v in enumerate(atuais) if v is not None]
    template_row = LINHA_JAN2000 + preenchidas[-1] if preenchidas else None
    print(f"  Linha template (formato): {template_row}")

    novos, revisoes, pendentes = [], [], {}

    for (ano, mes), valor in sorted(dados.items()):
        linha = row_of(ano, mes)
        if linha > fim:
            break
        i = linha - LINHA_JAN2000

        data_b = datas[i]
        if data_b is None:
            continue
        if isinstance(data_b, dt.datetime) and (data_b.year, data_b.month) != (ano, mes):
            raise RuntimeError(
                f"Divergência col B linha {linha}: esperado {mes:02d}/{ano}, "
                f"encontrado {data_b.month:02d}/{data_b.year}"
            )

        atual = atuais[i]
        if atual is None:
            novos.append(linha)
            pendentes[linha] = valor
        elif _mudou(atual, valor):
            revisoes.append((f"{mes:02d}/{ano}", atual, valor))
            pendentes[linha] = valor

    # ---- escrita ------------------------------------------------------
    for ini, f in blocos_contiguos(pendentes):
        for linha in range(ini, f + 1):
            sh.cell(linha, col_idx).value = pendentes[linha]
        print(f"  Escritas linhas {ini}:{f} ({f - ini + 1} meses)")

    # ---- formato da linha template -> linhas novas --------------------
    if novos and template_row is not None:
        src = sh.cell(template_row, col_idx)
        for ini, f in blocos_contiguos(novos):
            for linha in range(ini, f + 1):
                dst = sh.cell(linha, col_idx)
                dst.number_format = src.number_format
                dst.font      = copy(src.font)
                dst.fill      = copy(src.fill)
                dst.border    = copy(src.border)
                dst.alignment = copy(src.alignment)
            print(f"  Formato da linha {template_row} aplicado em {ini}:{f}")

    return {"novos": novos, "revisoes": revisoes, "template_row": template_row}


# ============================================================
# MAIN
# ============================================================
def main():
    # Baixa a planilha do Google Drive (em vez do drive montado em G:\).
    fid  = drive_io.find_file_id(XLSX_NAME)
    path = drive_io.download_to_temp(fid)
    print(f"[{dt.datetime.now():%Y-%m-%d %H:%M}] Atualizando PPI Transformadores...")

    print("\n[1/2] Baixando série do FRED...")
    dados = baixar_fred()
    print(f"  {len(dados)} meses com dados ({ANO_INICIO}+)")
    if not dados:
        print("Nenhum dado encontrado. Encerrando.")
        return

    print(f"\n[2/2] Abrindo {XLSX_NAME} via openpyxl...")
    wb = openpyxl.load_workbook(path, keep_vba=True, data_only=False)
    try:
        r = processar(wb[SHEET_NAME], dados)
        if r["novos"] or r["revisoes"]:
            wb.save(path)
            drive_io.update_file(fid, path)   # sobe de volta pro Drive
            print("Arquivo salvo.")
        else:
            print("Nada alterado — arquivo NÃO salvo.")
    finally:
        wb.close()

    print(f"\n== PPI Transformadores ==")
    print(f"Linhas novas : {len(r['novos'])}"
          + (f" ({r['novos'][0]}..{r['novos'][-1]})" if r["novos"] else ""))
    print(f"Revisões     : {len(r['revisoes'])}")
    for mes, antigo, novo in r["revisoes"]:
        print(f"  - {mes}: {antigo} → {novo}")


if __name__ == "__main__":
    main()
