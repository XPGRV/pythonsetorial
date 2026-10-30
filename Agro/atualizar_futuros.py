"""
atualizar_futuros.py — escreve os tickers e as formulas da aba Futuros.

Deixa a planilha PRONTA para o Bloomberg: cada bloco recebe a lista de
contratos do ciclo (ver tickers_futuros.py) e, ao lado, as formulas BDP/BDH
que o terminal vai preencher. Antes isso era feito na mao -- quando um
contrato vencia, alguem reescrevia a coluna e depois copiava os resultados
para a metade de cima.

Layout de cada bloco (colTk = coluna do ticker):
    colTk-1  data      =BDP(ticker,"LAST_TRADEABLE_DT")
    colTk    ticker    texto, escrito por este script
    colTk+1  atual     =BDP(ticker,"PX_LAST")
    colTk+2  1 sem.    =BDH(ticker,"PX_SETTLE",TODAY()-14,TODAY()-7,...)
    colTk+3  1 mes     =BDH(ticker,"PX_SETTLE",TODAY()-37,TODAY()-30,...)

So a METADE DE BAIXO (a das formulas) e' tocada. A de cima e' a que a
dashboard le e quem cuida dela e' o macro da planilha, no BeforeSave, depois
que o Bloomberg calcula.

Duas coisas que precisam continuar valendo:

* BDH e' formula MATRICIAL, mesmo devolvendo uma celula so ("Points=1"). Sem
  o t="array" a coleta para -- por isso sao escritas como ArrayFormula, e o
  drive_io preserva esse atributo no upload.

* O prefixo _xll. e' como o Excel guarda funcao de add-in. Mantido igual ao
  que ja estava na planilha.

Roda ANTES do Bloomberg, como o resto do pipeline.
"""

import argparse
import sys
from copy import copy
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.formula import ArrayFormula

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import drive_io

sys.path.insert(0, str(Path(__file__).resolve().parent))
from tickers_futuros import proximos

XLSX_NAME = "Agro.xlsm"
SHEET_NAME = "Futuros"
PRIMEIRA_LINHA = 5
ULTIMA_LINHA = 200

# (commodity, coluna do ticker)  B=2 soja, K=11 algodao, T=20 dolar, AC=29 milho
BLOCOS = [
    ("soja", 2),
    ("algodao", 11),
    ("dolar", 20),
    ("milho", 29),
]

FORMULAS = {
    -1: '=_xll.BDP({tk},"LAST_TRADEABLE_DT")',
    +1: '=_xll.BDP({tk},"PX_LAST")',
    +2: '=_xll.BDH({tk},"PX_SETTLE",TODAY()-14,TODAY()-7,"Dir=V","Fill=P","Points=1","Dates=H")',
    +3: '=_xll.BDH({tk},"PX_SETTLE",TODAY()-37,TODAY()-30,"Dir=V","Fill=P","Points=1","Dates=H")',
}


def eh_formula(valor):
    return isinstance(valor, ArrayFormula) or (
        isinstance(valor, str) and valor.startswith("=")
    )


def inicio_das_formulas(ws, col_tk):
    """Primeira linha da metade de baixo: tem ticker E formula no 'atual'."""
    for row in range(PRIMEIRA_LINHA, ULTIMA_LINHA + 1):
        if ws.cell(row, col_tk).value and eh_formula(ws.cell(row, col_tk + 1).value):
            return row
    return None


def linhas_ocupadas(ws, col_tk, inicio):
    """Quantas linhas o bloco de formulas ocupa hoje."""
    n = 0
    row = inicio
    while row <= ULTIMA_LINHA and ws.cell(row, col_tk).value:
        n += 1
        row += 1
    return n


def limpar_bloco(ws, col_tk, inicio):
    """Apaga a metade de formulas inteira, onde quer que ela esteja hoje."""
    n = 0
    row = inicio
    while row <= ULTIMA_LINHA and ws.cell(row, col_tk).value:
        for desloc in (-1, 0, 1, 2, 3):
            ws.cell(row, col_tk + desloc).value = None
        n += 1
        row += 1
    return n


def escrever_bloco(ws, col_tk, tickers, inicio, molde, dry_run=False):
    """Reescreve ticker + formulas do bloco. Devolve (escritas, limpas)."""
    escritas = limpas = 0
    tk_col_letra = get_column_letter(col_tk)

    for i, (_, _, ticker) in enumerate(tickers):
        row = inicio + i
        ref_tk = f"{tk_col_letra}{row}"

        if not dry_run:
            # ticker
            destino = ws.cell(row, col_tk)
            if destino.value != ticker:
                destino.value = ticker
            destino.number_format = ws.cell(molde, col_tk).number_format

            # formulas ao redor
            for desloc, gabarito in FORMULAS.items():
                cel = ws.cell(row, col_tk + desloc)
                texto = gabarito.format(tk=ref_tk)
                # BDH e' matricial de verdade; BDP tambem estava marcada assim
                # na planilha, entao mantemos o mesmo tratamento para as duas.
                cel.value = ArrayFormula(ref=cel.coordinate, text=texto)
                origem = ws.cell(molde, col_tk + desloc)
                cel.number_format = origem.number_format
                cel.font = copy(origem.font)
                cel.alignment = copy(origem.alignment)
                cel.border = copy(origem.border)
        escritas += 1

    # A linha de respiro, logo acima das formulas, tem de ficar SEM ticker.
    #
    # Quando o ciclo encurta, o bloco sobe (destino = PRIMEIRA_LINHA + n + 1) e
    # o respiro cai exatamente em cima da ultima linha da exibicao antiga. O
    # macro da planilha limpa a exibicao ate uma linha antes do respiro, entao
    # essa sobra ficava: um contrato repetido grudado no bloco novo -- foi o
    # CTZ7 em dobro no algodao.
    #
    # Para a dashboard isso nao e cosmetico. O parser le a coluna do ticker e
    # corta no primeiro vazio; com o respiro preenchido ele nao ve o corte,
    # segue para a metade das formulas e duplica a curva inteira.
    #
    # So do ticker para a direita: a coluna da data guarda um numero ali (100
    # no algodao, 112 na soja, 108 no milho) de origem desconhecida, que o
    # parser nao le. Fica onde esta.
    respiro = inicio - 1
    if not dry_run and respiro >= PRIMEIRA_LINHA:
        for desloc in (0, 1, 2, 3):
            ws.cell(respiro, col_tk + desloc).value = None

    # o ciclo pode ter encurtado: apaga o que sobrou embaixo
    row = inicio + len(tickers)
    while row <= ULTIMA_LINHA and ws.cell(row, col_tk).value:
        if not dry_run:
            for desloc in (-1, 0, 1, 2, 3):
                ws.cell(row, col_tk + desloc).value = None
        limpas += 1
        row += 1

    return escritas, limpas


def main():
    parser = argparse.ArgumentParser(
        description="Escreve tickers e formulas dos futuros na aba Futuros."
    )
    parser.add_argument("--dry-run", action="store_true",
                        help="Mostra o que faria, sem gravar.")
    args = parser.parse_args()

    file_id = drive_io.find_file_id(XLSX_NAME)
    path = drive_io.download_to_temp(file_id)

    wb = openpyxl.load_workbook(path, keep_vba=True, data_only=False)
    try:
        if SHEET_NAME not in wb.sheetnames:
            raise RuntimeError(
                f"Aba {SHEET_NAME!r} nao encontrada. Abas: {', '.join(wb.sheetnames)}"
            )
        ws = wb[SHEET_NAME]

        total = 0
        for commodity, col_tk in BLOCOS:
            inicio = inicio_das_formulas(ws, col_tk)
            if inicio is None:
                print(f"  {commodity:8s}: bloco sem formulas, pulado")
                continue

            antes = linhas_ocupadas(ws, col_tk, inicio)
            tickers = proximos(commodity)
            n = len(tickers)

            # A metade de baixo tem que caber DEPOIS da de cima. Como a de cima
            # e' espelho da de baixo, ela passa a ter n linhas -- entao as
            # formulas vao para 5 + n + 1 (uma linha de respiro no meio).
            # Sem isso, quando o ciclo cresce, o macro copiaria n linhas a
            # partir da 5 e escreveria por cima das proprias formulas: com 16
            # contratos de dolar comecando na linha 19, a copia iria ate a 20.
            destino = PRIMEIRA_LINHA + n + 1

            if not args.dry_run:
                limpar_bloco(ws, col_tk, inicio)          # tira do lugar antigo
            escritas, limpas = escrever_bloco(ws, col_tk, tickers, destino,
                                              molde=inicio, dry_run=args.dry_run)
            total += escritas + antes
            movida = "" if destino == inicio else f" (movida de {inicio})"
            print(f"  {commodity:8s}: formulas na linha {destino}{movida} | "
                  f"{antes} -> {n} contratos | "
                  f"{tickers[0][2]} ... {tickers[-1][2]}")

        prefix = "[DRY-RUN] " if args.dry_run else ""
        if total and not args.dry_run:
            wb.save(path)
        print(f"{prefix}Futuros: {total} linha(s) tocada(s).")
    finally:
        wb.close()
        archive = getattr(wb, "vba_archive", None)
        if archive is not None and getattr(archive, "fp", None) is not None:
            try:
                archive.close()
            except Exception:
                pass

    if total and not args.dry_run:
        drive_io.update_file(file_id, path)


if __name__ == "__main__":
    main()
