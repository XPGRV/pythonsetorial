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

import threading

import requests

BASE = "https://servicodados.ibge.gov.br/api/v3/agregados"
BASE_ANTIGA = "https://apisidra.ibge.gov.br/values"

# Qual endereco tentar primeiro. Reordenado em tempo de execucao: o
# primeiro que responder vira o preferido do processo.
_ORDEM = ["apisidra", "servicodados"]

# Prazo de parede, somado sobre conectar + ler. O timeout do requests sozinho
# nao basta: ele e por operacao de socket, e um servidor que entrega um byte
# de vez em quando nunca o dispara. Foi assim que tres jobs da coleta ficaram
# 20 minutos mudos em 28/10/2026, ate o GitHub mata-los por timeout-minutes.
PRAZO_PADRAO = 120


def buscar_json(url: str, params: dict | None = None, timeout: int = 60,
                prazo: int | None = None):
    """GET que devolve JSON e desiste no prazo, aconteca o que acontecer.

    A requisicao roda numa thread daemon; quem conta o tempo e' quem chama.
    Se estourar, a thread continua pendurada no socket, mas por ser daemon nao
    impede o processo de encerrar.
    """
    prazo = prazo or PRAZO_PADRAO
    resultado: dict = {}

    def tentar():
        try:
            r = requests.get(url, params=params, timeout=(10, timeout))
            r.raise_for_status()
            resultado["dados"] = r.json()
        except Exception as exc:          # guardado para o chamador relancar
            resultado["erro"] = exc

    thread = threading.Thread(target=tentar, daemon=True)
    thread.start()
    thread.join(prazo)

    if thread.is_alive():
        raise RuntimeError(
            f"O IBGE nao respondeu em {prazo}s: {url}. A requisicao foi "
            "abandonada para o job nao ficar pendurado."
        )
    if "erro" in resultado:
        raise resultado["erro"]
    return resultado["dados"]


def _via_apisidra(tabela, variaveis, periodos, classificacoes, localidades, timeout):
    """Endereco antigo. A resposta JA vem no formato de saida; so tiramos o
    cabecalho descritivo, que e sempre a primeira linha."""
    if localidades != "N1[all]":
        raise RuntimeError(
            f"apisidra: nao sei montar a URL para localidades={localidades!r}"
        )
    caminho = (
        f"{BASE_ANTIGA}/t/{tabela}/n1/1"
        f"/v/{','.join(variaveis)}/p/{','.join(periodos)}"
    )
    for cid, cats in classificacoes:
        caminho += f"/c{cid}/{','.join(cats)}"

    dados = buscar_json(caminho, timeout=timeout)
    if not isinstance(dados, list) or not dados:
        return []
    return dados[1:]


def _via_servicodados(tabela, variaveis, periodos, classificacoes, localidades, timeout):
    """Endereco novo. Resposta aninhada, convertida para o formato de saida."""
    url = f"{BASE}/{tabela}/periodos/{'|'.join(periodos)}/variaveis/{'|'.join(variaveis)}"
    params = {
        "localidades": localidades,
        "classificacao": "|".join(f"{cid}[{','.join(cats)}]" for cid, cats in classificacoes),
    }
    dados = buscar_json(url, params=params, timeout=timeout)

    linhas: list[dict] = []
    for variavel in dados:
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


_CAMINHOS = {"apisidra": _via_apisidra, "servicodados": _via_servicodados}


def consultar(
    tabela: str,
    variaveis: list[str],
    periodos: list[str],
    classificacoes: list[tuple[str, list[str]]],
    localidades: str = "N1[all]",
    timeout: int = 90,
) -> list[dict]:
    """Linhas da tabela, pelo endereco do IBGE que estiver de pe.

    Tenta os dois; so levanta erro se nenhum responder. Lista vazia continua
    significando "o IBGE respondeu, mas nao ha dado publicado para o periodo".
    """
    falhas = []
    for nome in list(_ORDEM):
        try:
            linhas = _CAMINHOS[nome](
                tabela, variaveis, periodos, classificacoes, localidades, timeout
            )
        except Exception as exc:
            falhas.append(f"{nome}: {type(exc).__name__} {exc}")
            continue

        if nome != _ORDEM[0]:
            # Quem respondeu passa a ser o primeiro da fila: as proximas
            # chamadas deste processo nao pagam de novo pelo que esta fora.
            _ORDEM.remove(nome)
            _ORDEM.insert(0, nome)
            print(f"  [IBGE] usando {nome} (o outro endereco nao respondeu)", flush=True)
        return linhas

    raise RuntimeError(
        "Nenhum endereco do IBGE respondeu para a tabela "
        f"{tabela}, periodos {','.join(periodos)}. " + " | ".join(falhas)
    )
