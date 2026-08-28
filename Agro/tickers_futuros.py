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

# prefixo, meses do ciclo, sufixo, digitos do ano, separador, defasagem, limite
#
# defasagem = quantos meses antes do codigo o contrato deixa de negociar.
#   0  CBOT/ICE e o milho da B3 vencem no meio do mes do proprio codigo
#      (S Q6 -> 14/08/26, CTV6 -> 08/10/26, CRDU6 -> 15/09/26).
#   1  o dolar da B3 negocia ate o ultimo dia util do mes ANTERIOR
#      (UCQ26, codigo de agosto, vence 31/07/26) -- por isso ele sai do ciclo
#      um mes antes dos outros.
#
# limite = ate onde ir. A lista NAO tem tamanho fixo: ela encolhe conforme os
# contratos vencem e da' um salto quando o vencimento de referencia passa.
#   ("terminal", mes, n)  vai ate a n-esima ocorrencia daquele mes.
#                         Soja e milho fecham o ciclo em Novembro, algodao em
#                         Dezembro; manter 2 e' o mesmo criterio que a janela
#                         do grafico de futuros da dashboard usa.
#   ("fim_ano", k)        vai ate dezembro do ano corrente + k (dolar: negocia
#                         todo mes, entao o corte natural e' o fim do ano).
CONTRATOS = {
    "soja":    ("S",   (1, 3, 5, 7, 8, 9, 11), "Comdty", 1, " ", 0, ("terminal", 11, 2)),
    "algodao": ("CT",  (3, 5, 7, 10, 12),      "Comdty", 1, "",  0, ("terminal", 12, 2)),
    "milho":   ("CRD", (1, 3, 5, 7, 9, 11),    "Comdty", 1, "",  0, ("terminal", 11, 2)),
    "dolar":   ("UC",  tuple(range(1, 13)),    "Curncy", 2, "",  1, ("fim_ano", 1)),
}


def ticker(commodity: str, ano: int, mes: int) -> str:
    prefixo, _, sufixo, digitos, sep, _, _ = CONTRATOS[commodity]
    ano_txt = str(ano)[-digitos:] if digitos == 2 else str(ano % 10)
    return f"{prefixo}{sep}{CODIGO_MES[mes]}{ano_txt} {sufixo}"


def proximos(commodity: str, hoje: date | None = None):
    """Os contratos que ainda negociam, do mais proximo ate o limite do ciclo.

    Devolve (ano, mes, ticker) -- ano/mes sao os do CODIGO do contrato, que e'
    o que a dashboard usa como vencimento no eixo X.
    """
    hoje = hoje or date.today()
    _, ciclo, _, _, _, defasagem, limite = CONTRATOS[commodity]

    # Um contrato ainda vale enquanto o mes de negociacao (codigo - defasagem)
    # nao ficou para tras. No mes corrente ele continua na lista: so vence no
    # meio/fim do mes, e a planilha o mantem ate virar.
    ordinal = hoje.year * 12 + hoje.month - 1 + defasagem

    saida = []
    if limite[0] == "terminal":
        _, mes_terminal, quantos = limite
        vistos = 0
        while vistos < quantos:
            ano, mes = divmod(ordinal, 12)
            mes += 1
            if mes in ciclo:
                saida.append((ano, mes, ticker(commodity, ano, mes)))
                if mes == mes_terminal:
                    vistos += 1
            ordinal += 1
    else:  # ("fim_ano", k) — ate dezembro do ano corrente + k
        ano_limite = hoje.year + limite[1]
        while True:
            ano, mes = divmod(ordinal, 12)
            mes += 1
            if ano > ano_limite:
                break
            if mes in ciclo:
                saida.append((ano, mes, ticker(commodity, ano, mes)))
            ordinal += 1
    return saida


def main():
    parser = argparse.ArgumentParser(description="Gera os tickers dos futuros.")
    parser.add_argument("--data", default=None,
                        help="simula outra data (AAAA-MM-DD), p/ conferir a virada do ciclo")
    args = parser.parse_args()
    hoje = date.fromisoformat(args.data) if args.data else None
    for commodity in CONTRATOS:
        lista = proximos(commodity, hoje=hoje)
        print(f"\n{commodity.upper()} ({len(lista)} contratos)")
        for ano, mes, tk in lista:
            print(f"  {mes:02d}/{ano}  {tk}")


if __name__ == "__main__":
    main()
