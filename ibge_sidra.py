"""Consulta tabelas do SIDRA pela API oficial servicodados do IBGE.

Por que existe
--------------
Os coletores liam o SIDRA por apisidra.ibge.gov.br. Entre 10/09 e 15/09/2026
o IBGE colocou esse endereco atras do Cloudflare, que responde 403 ("Just a
moment...") a qualquer cliente que nao seja navegador -- e o bloqueio nao cai
trocando o User-Agent: com o mesmo cabecalho o curl passa e o requests nao.

As mesmas tabelas saem pela API v3 do servicodados.ibge.gov.br, que e a API
documentada do IBGE para uso programatico e nao tem esse bloqueio (o IPCA do
FrangoBR e o CarRental ja usavam ela).

Formato de saida
----------------
Devolve as linhas no formato da API antiga (sem a linha de cabecalho), para os
scripts nao precisarem mudar o parse:

    {"D1C": localidade, "D2C": variavel, "D3C": periodo,
     "D4C": 1a classificacao, "D5C": 2a, ..., "V": valor em texto}

As classificacoes ocupam D4C, D5C, ... na ordem em que forem passadas --
a mesma ordem em que apareciam na URL antiga. O valor V vem como texto,
inclusive os marcadores "X", "-" e "...", igual antes.
"""
from __future__ import annotations

import requests

BASE = "https://servicodados.ibge.gov.br/api/v3/agregados"


def consultar(
    tabela: str,
    variaveis: list[str],
    periodos: list[str],
    classificacoes: list[tuple[str, list[str]]],
    localidades: str = "N1[all]",
    timeout: int = 90,
) -> list[dict]:
    """Linhas da tabela no formato da apisidra. Lista vazia = sem dado publicado."""
    url = f"{BASE}/{tabela}/periodos/{'|'.join(periodos)}/variaveis/{'|'.join(variaveis)}"
    params = {
        "localidades": localidades,
        "classificacao": "|".join(f"{cid}[{','.join(cats)}]" for cid, cats in classificacoes),
    }
    r = requests.get(url, params=params, timeout=timeout)
    r.raise_for_status()

    linhas: list[dict] = []
    for variavel in r.json():
        for resultado in variavel.get("resultados", []):
            categoria = {
                c["id"]: next(iter(c["categoria"]))
                for c in resultado.get("classificacoes", [])
            }
            for serie in resultado.get("series", []):
                local = serie.get("localidade", {}).get("id")
                for periodo, valor in serie.get("serie", {}).items():
                    linha = {"D1C": local, "D2C": variavel["id"], "D3C": periodo, "V": valor}
                    for i, (cid, _) in enumerate(classificacoes, start=4):
                        linha[f"D{i}C"] = categoria.get(cid)
                    linhas.append(linha)
    return linhas
