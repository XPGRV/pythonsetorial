"""
tickers_futuros.py — gera os tickers dos contratos futuros da aba Futuros.

Hoje a lista e' escrita na mao: quando um contrato vence, ele sai do ciclo e
alguem precisa reescrever a coluna inteira. Os ciclos sao deterministicos, so
que cada bolsa tem o seu, entao da' para gerar.

    Soja     S    CBOT   F H K N Q U X     Jan Mar Mai Jul Ago Set Nov
    Algodao  CT   ICE      H K N V Z       Mar Mai Jul Out Dez
    Milho    CRD  B3     F H K N U X       Jan Mar Mai Jul Set Nov
    Dolar    UC   B3     todos os meses

O sufixo tambem muda: as commodities sao "Comdty" e o dolar e' "Curncy". E o
ano vem com 1 digito nas commodities ("S Q6") e 2 no dolar ("UCQ26") -- foi o
que exigiu aceitar os dois formatos no parser da dashboard.

Uso:
    python tickers_futuros.py            # imprime os tickers de hoje
    python tickers_futuros.py --meses 30 # horizonte maior
"""

import argparse
from datetime import date

# Codigo Bloomberg do mes de vencimento.
CODIGO_MES = {
    1: "F", 2: "G", 3: "H", 4: "J", 5: "K", 6: "M",
    7: "N", 8: "Q", 9: "U", 10: "V", 11: "X", 12: "Z",
}

# prefixo, meses do ciclo, sufixo, digitos do ano, separador, defasagem, quantidade
#
# defasagem = quantos meses antes do codigo o contrato deixa de negociar.
#   0  CBOT/ICE e o milho da B3 vencem no meio do mes do proprio codigo
#      (S Q6 -> 14/08/26, CTV6 -> 08/10/26, CRDU6 -> 15/09/26).
#   1  o dolar da B3 negocia ate o ultimo dia util do mes ANTERIOR
#      (UCQ26, codigo de agosto, vence 31/07/26) -- por isso ele sai do ciclo
#      um mes antes dos outros.
#
# quantidade = quantos contratos manter na planilha, espelhando o que ja estava
# la' na mao. O dolar tem contrato todo mes, entao um horizonte em meses daria
# uma lista longa demais; melhor contar contratos.
CONTRATOS = {
    "soja":    ("S",   (1, 3, 5, 7, 8, 9, 11), "Comdty", 1, " ", 0, 17),
    "algodao": ("CT",  (3, 5, 7, 10, 12),      "Comdty", 1, "",  0, 12),
    "milho":   ("CRD", (1, 3, 5, 7, 9, 11),    "Comdty", 1, "",  0, 8),
    "dolar":   ("UC",  tuple(range(1, 13)),    "Curncy", 2, "",  1, 13),
}


def ticker(commodity: str, ano: int, mes: int) -> str:
    prefixo, _, sufixo, digitos, sep, _, _ = CONTRATOS[commodity]
    ano_txt = str(ano)[-digitos:] if digitos == 2 else str(ano % 10)
    return f"{prefixo}{sep}{CODIGO_MES[mes]}{ano_txt} {sufixo}"


def proximos(commodity: str, quantidade: int | None = None, hoje: date | None = None):
    """Os proximos contratos que ainda negociam, do mais proximo ao mais longo.

    Devolve (ano, mes, ticker) -- ano/mes sao os do CODIGO do contrato, que e'
    o que a dashboard usa como vencimento no eixo X.
    """
    hoje = hoje or date.today()
    _, ciclo, _, _, _, defasagem, padrao = CONTRATOS[commodity]
    quantidade = padrao if quantidade is None else quantidade
    # Um contrato ainda vale enquanto o mes de negociacao (codigo - defasagem)
    # nao ficou para tras. No mes corrente ele continua na lista: so vence no
    # meio/fim do mes, e a planilha o mantem ate virar.
    atual = hoje.year * 12 + hoje.month - 1
    saida, ordinal = [], atual + defasagem
    while len(saida) < quantidade:
        ano, mes = divmod(ordinal, 12)
        mes += 1
        if mes in ciclo:
            saida.append((ano, mes, ticker(commodity, ano, mes)))
        ordinal += 1
    return saida


def main():
    parser = argparse.ArgumentParser(description="Gera os tickers dos futuros.")
    parser.add_argument("--quantidade", type=int, default=None,
                        help="quantos contratos por commodity (padrao: o de CONTRATOS)")
    args = parser.parse_args()
    for commodity in CONTRATOS:
        lista = proximos(commodity, args.quantidade)
        print(f"\n{commodity.upper()} ({len(lista)} contratos)")
        for ano, mes, tk in lista:
            print(f"  {mes:02d}/{ano}  {tk}")


if __name__ == "__main__":
    main()
