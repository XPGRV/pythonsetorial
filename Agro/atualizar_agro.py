"""
atualizar_agro.py — consolidador do Agro (aba BBG_Dados).

O Bloomberg traz os precos (colunas E..X, valores ja congelados pelo macro),
mas as colunas de DESCONTO sao formulas que alguem precisa "arrastar" para
baixo. Sem isso elas ficam para tras e os graficos de desconto da dashboard
(algodao, soja e milho) param no dia em que o ultimo arraste manual parou.

    I = IFERROR(G-E,"-")   J = IFERROR(I/E,"-")   Algodao  (Barreiras - CBOT)
    Q = IFERROR(M-K,"-")   R = IFERROR(Q/K,"-")   Soja     (Paranagua - CBOT)
    Y = IFERROR(U-S,"-")   Z = IFERROR(Y/S,"-")   Milho    (Campinas  - CBOT)

Sao formulas 100% relativas e sempre na propria linha, entao "arrastar" e so
copiar a ultima linha que tem formula para as linhas novas.

ATE ONDE ESTENDER — so ate a ultima linha com preco de verdade. Nas linhas
futuras (o calendario da coluna D vai ate 2030) as fontes estao VAZIAS, e
vazio-vazio = 0 no Excel: o IFERROR nao pega, e a dashboard leria um desconto
zero falso todo dia ate 2030. Nas linhas antigas sem cotacao isso nao acontece
porque o Bloomberg deixa "#N/A N/A", que da erro e cai no "-".

CACHE / ORDEM DE EXECUCAO — ao salvar, o openpyxl apaga o valor em cache de
TODAS as formulas do arquivo. A dashboard le com SheetJS, que so le cache e nao
recalcula. Como aqui as unicas formulas relevantes sao os proprios descontos,
rodar sem recalcular deixaria os tres graficos de desconto em branco ate o
Excel abrir a planilha de novo. Duas saidas:

  * rodar ANTES do passo Bloomberg — o RefreshAll+F9 de la restaura o cache
    (mesma ordem dos outros pipelines); custa 1 dia de atraso na linha nova; ou
  * rodar com --recalc, que abre o Excel local so para recalcular e salvar
    (sem atraso, mas exige Excel na maquina).

Uso:
    python atualizar_agro.py --dry-run     # mostra o que faria
    python atualizar_agro.py               # grava as formulas
    python atualizar_agro.py --recalc      # grava e recalcula no Excel local
"""

import argparse
import re
import time
from copy import copy
from pathlib import Path

import openpyxl

WORKBOOK_PATH = Path(
    r"G:\.shortcut-targets-by-id\11q1ngdqbySnDMCQSP4fiwPztT6ju2F4Z"
    r"\Arquivos\Setorial - Database\Agro.xlsm"
)
SHEET_NAME = "BBG_Dados"
FIRST_DATA_ROW = 4

# (nome, colunas de destino, colunas-fonte que definem ate onde ha dado)
# As fontes sao o par CBOT + praca local de cada commodity; o alvo e a ultima
# linha em que qualquer uma das duas tem numero.
GRUPOS = [
    ("Algodao", (9, 10), (5, 7)),     # I,J  <- E (CBOT), G (Barreiras)
    ("Soja",    (17, 18), (11, 13)),  # Q,R  <- K (CBOT), M (Paranagua)
    ("Milho",   (25, 26), (19, 21)),  # Y,Z  <- S (CBOT), U (Campinas)
]

REF = re.compile(r"(?<![A-Za-z0-9_])(\$?[A-Za-z]{1,3})(\$?)(\d+)")


def shift(value, delta):
    """Desloca as referencias de linha de uma formula, respeitando os '$'."""
    if not isinstance(value, str) or not value.startswith("="):
        return value
    return REF.sub(
        lambda m: m.group(0)
        if m.group(2) == "$"
        else f"{m.group(1)}{m.group(2)}{int(m.group(3)) + delta}",
        value,
    )


def copy_cell(origem, destino, delta):
    destino.value = shift(origem.value, delta)
    destino._style = copy(origem._style)
    destino.number_format = origem.number_format
    destino.font = copy(origem.font)
    destino.fill = copy(origem.fill)
    destino.border = copy(origem.border)
    destino.alignment = copy(origem.alignment)


def last_formula(ws, col, start=FIRST_DATA_ROW):
    """Ultima linha da coluna que ainda tem formula — serve de molde."""
    for row in range(ws.max_row, start - 1, -1):
        valor = ws.cell(row, col).value
        if isinstance(valor, str) and valor.startswith("="):
            return row
    return None


def last_numeric(ws, col, start=FIRST_DATA_ROW):
    """Ultima linha com numero de verdade (texto de erro e vazio nao contam)."""
    ultima = 0
    for row in range(start, ws.max_row + 1):
        if isinstance(ws.cell(row, col).value, (int, float)):
            ultima = row
    return ultima


def extend(ws, molde, alvo, colunas, dry_run=False):
    if not molde or alvo <= molde:
        return 0
    changes = 0
    for row in range(molde + 1, alvo + 1):
        for col in colunas:
            # So completa lacuna; nunca substitui o que ja esta escrito.
            if ws.cell(row, col).value not in (None, ""):
                continue
            if not dry_run:
                copy_cell(ws.cell(molde, col), ws.cell(row, col), row - molde)
            changes += 1
    return changes


def recalcular_com_excel(path, espera_max=20):
    """Abre no Excel local so para recalcular e regravar o cache das formulas.

    NAO esperar CalculationState virar xlDone: fora do PC Bloomberg os links BDH
    nao resolvem e o Excel fica em xlPending para sempre. O Calculate() ja
    resolve o que da para resolver de forma sincrona (validado: as formulas de
    desconto saem com o valor certo), entao a espera aqui e so uma folga.
    """
    import win32com.client as win32

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AutomationSecurity = 3   # forca macros desabilitadas nesta sessao
    excel.EnableEvents = False     # nao dispara Workbook_Open/Auto_Open
    excel.AskToUpdateLinks = False
    try:
        # UpdateLinks=0: nao tenta atualizar os links BDH (so existem no PC Bloomberg)
        wb = excel.Workbooks.Open(str(path), UpdateLinks=0)
        excel.Calculate()
        esperou = 0.0
        while excel.CalculationState != 0 and esperou < espera_max:  # 0 = xlDone
            time.sleep(0.5)
            esperou += 0.5
        wb.Save()
        wb.Close(SaveChanges=False)
    finally:
        excel.Quit()


def close_workbook(wb):
    """Fecha o workbook e o vba_archive (que o openpyxl deixa aberto)."""
    wb.close()
    archive = getattr(wb, "vba_archive", None)
    if archive is not None and getattr(archive, "fp", None) is not None:
        try:
            archive.close()
        except Exception:
            pass


def main():
    parser = argparse.ArgumentParser(
        description="Arrasta as formulas de desconto da aba BBG_Dados do Agro.xlsm."
    )
    parser.add_argument("--dry-run", action="store_true",
                        help="Mostra o que seria estendido, sem gravar.")
    parser.add_argument("--recalc", action="store_true",
                        help="Apos gravar, abre o Excel local para recalcular "
                             "e restaurar o cache das formulas.")
    args = parser.parse_args()

    if not WORKBOOK_PATH.exists():
        raise FileNotFoundError(f"Planilha nao encontrada: {WORKBOOK_PATH}")

    wb = openpyxl.load_workbook(WORKBOOK_PATH, keep_vba=True, data_only=False)
    try:
        if SHEET_NAME not in wb.sheetnames:
            raise RuntimeError(
                f"Aba {SHEET_NAME!r} nao encontrada. Abas: {', '.join(wb.sheetnames)}"
            )
        ws = wb[SHEET_NAME]

        total = 0
        for nome, destino, fontes in GRUPOS:
            alvo = max(last_numeric(ws, c) for c in fontes)
            molde = last_formula(ws, destino[0])
            changes = extend(ws, molde, alvo, destino, dry_run=args.dry_run)
            total += changes
            faltando = max(0, alvo - molde) if molde else 0
            print(f"  {nome:8s}: formulas ate {molde} | precos ate {alvo} | "
                  f"{faltando} linha(s) a preencher | {changes} celula(s)")

        prefix = "[DRY-RUN] " if args.dry_run else ""
        if total and not args.dry_run:
            wb.save(WORKBOOK_PATH)
        print(f"{prefix}Consolidador Agro: {total} celula(s) estendida(s).")
    finally:
        close_workbook(wb)

    if total and not args.dry_run:
        if args.recalc:
            print("Recalculando no Excel para restaurar o cache das formulas...")
            recalcular_com_excel(WORKBOOK_PATH)
            print("Recalculo concluido.")
        else:
            print("ATENCAO: as formulas ficaram sem valor em cache. A dashboard so "
                  "mostrara os descontos depois que o Excel recalcular (passo "
                  "Bloomberg ou --recalc).")


if __name__ == "__main__":
    main()
