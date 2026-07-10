# -*- coding: utf-8 -*-
"""
Atualizacao mensal IPCA — aba 'Preco Carros' de CarRental.xlsm

Escreve nas colunas:
  C = Automovel Novo   — variacao mensal %MoM (IPCA subitem 5102001, cat 315[7641])
  D = Automovel Usado  — variacao mensal %MoM (IPCA subitem 5102020)

A serie do IPCA tem quebras metodologicas, entao ela e emendada a partir de
quatro tabelas do SIDRA (API v3). O codigo da categoria do 'usado' muda entre
as tabelas antigas (7651) e as novas (107654); o 'novo' e sempre 7641.

  Tabela 655  : jan/2000 -> jun/2006   usado = 7651
  Tabela 2938 : jul/2006 -> dez/2011   usado = 107654
  Tabela 1419 : jan/2012 -> dez/2019   usado = 107654
  Tabela 7060 : jan/2020 -> presente   usado = 107654  (periodos buscados dinamicamente)

Variavel:
  63 = variacao mensal (%)

Os valores sao gravados como fracao (ex.: variacao de 1,39% -> 0.0139) e as
celulas recebem formato percentual '0.00%', exibindo '1,39%'.

Escrita via xlwings — o arquivo tem VBA e formulas de array do Bloomberg na aba
'Peers', que o openpyxl poderia corromper. O xlwings dirige o Excel real e
preserva formatacao, formulas e demais abas intactas.
"""

import time
import datetime as dt
from pathlib import Path

import requests
import xlwings as xw


# ============================================================
# CONFIG
# ============================================================
DRIVE_ROOT = Path(
    r"G:\.shortcut-targets-by-id\11q1ngdqbySnDMCQSP4fiwPztT6ju2F4Z"
)
XLSX_NAME  = "CarRental.xlsm"
SHEET_NAME = "Preço Carros"

COL_NOVO  = "C"   # Automovel Novo  — %MoM
COL_USADO = "D"   # Automovel Usado — %MoM
COL_DATA  = "B"   # datas mensais — usada para validacao

ANO_BASE = 2000
ROW_BASE = 5      # linha 5 = jan/2000

VARIAVEL   = "63"       # variacao mensal (%)
# Formato aplicado as celulas: fracao exibida como percentual com 2 casas.
# Usa-se NumberFormatLocal com a string pt-BR (virgula decimal) porque este
# Excel roda em locale 416; passar "0.00%" faria o Excel ler o "." como
# separador de milhar e gravar "#.000%" (sem casas decimais).
FORMATO_PCT = "0,00%"

# Tabelas do SIDRA emendadas (ordem crescente no tempo).
# Cada entrada mapeia codigo da categoria (classificacao 315) -> serie interna.
AGREGADOS = [
    {
        "id": 655,
        "label": "jan/2000–jun/2006",
        "ini": (2000, 1),
        "fim": (2006, 6),
        "codigos": {"7641": "novo", "7651": "usado"},
    },
    {
        "id": 2938,
        "label": "jul/2006–dez/2011",
        "ini": (2006, 7),
        "fim": (2011, 12),
        "codigos": {"7641": "novo", "107654": "usado"},
    },
    {
        "id": 1419,
        "label": "jan/2012–dez/2019",
        "ini": (2012, 1),
        "fim": (2019, 12),
        "codigos": {"7641": "novo", "107654": "usado"},
    },
    {
        "id": 7060,
        "label": "jan/2020–presente",
        "ini": (2020, 1),
        "fim": None,   # dinamico: buscado via /periodos
        "codigos": {"7641": "novo", "107654": "usado"},
    },
]

BASE_URL    = "https://servicodados.ibge.gov.br/api/v3/agregados"
TIMEOUT     = 60
MAX_RETRIES = 3
RETRY_SLEEP = 5
EPS         = 1e-6   # tolerancia p/ considerar revisao de valor


# ============================================================
# HELPERS
# ============================================================
def localizar_planilha() -> Path:
    """Localiza a planilha sem depender dos nomes das pastas intermediarias."""
    encontrados = list(DRIVE_ROOT.rglob(XLSX_NAME))
    if not encontrados:
        raise FileNotFoundError(
            f"Planilha {XLSX_NAME!r} nao encontrada dentro de {DRIVE_ROOT}"
        )
    if len(encontrados) > 1:
        caminhos = "\n".join(f"  - {p}" for p in encontrados)
        raise RuntimeError(
            f"Foram encontradas varias planilhas chamadas {XLSX_NAME!r}:\n{caminhos}"
        )
    return encontrados[0]


def row_of(ano: int, mes: int) -> int:
    """Linha 5 = jan/2000."""
    return ROW_BASE + (ano - ANO_BASE) * 12 + (mes - 1)


def meses_entre(ini: tuple[int, int], fim: tuple[int, int]) -> list[str]:
    """Lista 'AAAAMM' de ini ate fim (inclusive)."""
    (ay, am), (by, bm) = ini, fim
    out, y, m = [], ay, am
    while (y, m) <= (by, bm):
        out.append(f"{y}{m:02d}")
        m += 1
        if m == 13:
            m, y = 1, y + 1
    return out


def parse_valor(raw):
    """Converte valor do SIDRA em float; ausencias ('...', '-', '') viram None."""
    if raw is None:
        return None
    s = str(raw).strip()
    if s in ("...", "-", "", "X", ".."):
        return None
    try:
        return float(s.replace(",", "."))
    except ValueError:
        return None


def _get(url: str) -> list:
    last_err = None
    for tentativa in range(1, MAX_RETRIES + 1):
        try:
            r = requests.get(url, timeout=TIMEOUT)
            r.raise_for_status()
            return r.json()
        except Exception as e:
            last_err = e
            if tentativa < MAX_RETRIES:
                time.sleep(RETRY_SLEEP)
    raise RuntimeError(f"API SIDRA falhou em {MAX_RETRIES} tentativas: {last_err}")


def periodos_disponiveis(agregado_id: int) -> list[str]:
    """Consulta os periodos publicados para um agregado (para o 7060, dinamico)."""
    data = _get(f"{BASE_URL}/{agregado_id}/periodos")
    return [p["id"] for p in data]


def fetch_agregado(agg: dict) -> dict:
    """Retorna {(ano, mes): {'novo': v, 'usado': v}} para um agregado."""
    codigos = agg["codigos"]

    if agg["fim"] is None:
        disp = periodos_disponiveis(agg["id"])
        ini = f"{agg['ini'][0]}{agg['ini'][1]:02d}"
        periodos = [p for p in disp if p >= ini]
        print(f"    {agg['id']}: {len(periodos)} periodos ({periodos[0]} -> {periodos[-1]})")
    else:
        periodos = meses_entre(agg["ini"], agg["fim"])

    url = (
        f"{BASE_URL}/{agg['id']}"
        f"/periodos/{'|'.join(periodos)}"
        f"/variaveis/{VARIAVEL}"
        f"?localidades=N1[all]&classificacao=315[{','.join(codigos)}]"
    )
    data = _get(url)

    janela_ini = agg["ini"]
    janela_fim = agg["fim"]

    resultado: dict = {}
    for bloco in data:
        for res in bloco.get("resultados", []):
            cat = res.get("classificacoes", [{}])[0].get("categoria", {})
            codigo = next(iter(cat), None)
            nome = codigos.get(str(codigo))
            if not nome:
                continue
            serie = res.get("series", [{}])[0].get("serie", {})
            for periodo, valor_str in serie.items():
                ano, mes = int(periodo[:4]), int(periodo[4:])
                if (ano, mes) < janela_ini:
                    continue
                if janela_fim is not None and (ano, mes) > janela_fim:
                    continue
                valor = parse_valor(valor_str)
                if valor is None:
                    continue
                resultado.setdefault((ano, mes), {})[nome] = valor
    return resultado


# ============================================================
# MAIN
# ============================================================
def main():
    path = localizar_planilha()
    print(f"[{dt.datetime.now():%Y-%m-%d %H:%M}] Abrindo {path.name} via xlwings...")

    app = xw.App(visible=False, add_book=False)
    try:
        app.api.AskToUpdateLinks = False
        app.display_alerts       = False
        app.screen_updating      = False

        wb = app.books.open(str(path), update_links=0, read_only=False)
        try:
            resultado = processar(wb)
            if resultado["salvar"]:
                wb.api.SaveLinkValues = True
                wb.save()
                print("Arquivo salvo.")
            else:
                print("Nada alterado — arquivo NAO salvo.")
        finally:
            wb.close()
    finally:
        app.quit()

    print_relatorio(resultado)


# ============================================================
# PROCESSAMENTO
# ============================================================
def processar(wb) -> dict:
    sh = wb.sheets[SHEET_NAME]

    last_row  = sh.range(f"{COL_DATA}{ROW_BASE}").end("down").row
    novo_vals = sh.range(f"{COL_NOVO}{ROW_BASE}:{COL_NOVO}{last_row}").value
    usado_vals = sh.range(f"{COL_USADO}{ROW_BASE}:{COL_USADO}{last_row}").value
    datas_col = sh.range(f"{COL_DATA}{ROW_BASE}:{COL_DATA}{last_row}").value

    def as_list(v):
        return v if isinstance(v, list) else [v]

    novo_vals  = as_list(novo_vals)
    usado_vals = as_list(usado_vals)
    datas_col  = as_list(datas_col)

    # ---- buscar e consolidar ----------------------------------------
    # Busca sempre as 4 tabelas (poucas chamadas, segundos). As tabelas
    # historicas sao fechadas: em regime permanente o diff nao gera escrita
    # nelas, entao a atualizacao continua "incremental" na pratica — so meses
    # novos ou revisados chegam a ser gravados. Buscar tudo torna o script
    # auto-corretivo (ex.: mudanca de unidade/formato reescreve a serie toda).
    print("Consultando SIDRA (4 tabelas emendadas)...")
    dados: dict = {}
    for agg in AGREGADOS:
        parcial = fetch_agregado(agg)
        for chave, vals in parcial.items():
            dados.setdefault(chave, {}).update(vals)
        print(f"  tabela {agg['id']} ({agg['label']}): {len(parcial)} meses")

    # ---- comparar e preparar escrita --------------------------------
    contadores = {"novas_novo": 0, "rev_novo": 0, "novas_usado": 0, "rev_usado": 0}
    revisoes_novo, revisoes_usado, ausentes = [], [], []
    escrita_pendente: dict[int, list] = {}

    for (ano, mes), vals in sorted(dados.items()):
        linha = row_of(ano, mes)
        idx = linha - ROW_BASE
        rotulo = f"{mes:02d}/{ano}"

        if idx < 0 or idx >= len(datas_col):
            ausentes.append(rotulo)   # fora do range de datas da planilha
            continue

        # validacao: a data na coluna B tem que bater com o mes esperado
        data_b = datas_col[idx]
        if isinstance(data_b, dt.datetime) and (data_b.year != ano or data_b.month != mes):
            raise RuntimeError(
                f"Divergencia col {COL_DATA} linha {linha}: esperado {rotulo}, "
                f"encontrado {data_b.month:02d}/{data_b.year}"
            )

        atual_n = novo_vals[idx]  if idx < len(novo_vals)  else None
        atual_u = usado_vals[idx] if idx < len(usado_vals) else None
        # API devolve a variacao em % (ex.: 1.39); grava-se como fracao (0.0139)
        novo_n  = vals.get("novo")
        novo_u  = vals.get("usado")
        if novo_n is not None:
            novo_n /= 100
        if novo_u is not None:
            novo_u /= 100

        linha_mudou = False

        if novo_n is not None:
            if atual_n is None:
                contadores["novas_novo"] += 1
                linha_mudou = True
            elif abs(float(atual_n) - novo_n) > EPS:
                contadores["rev_novo"] += 1
                revisoes_novo.append((rotulo, atual_n, novo_n))
                linha_mudou = True

        if novo_u is not None:
            if atual_u is None:
                contadores["novas_usado"] += 1
                linha_mudou = True
            elif abs(float(atual_u) - novo_u) > EPS:
                contadores["rev_usado"] += 1
                revisoes_usado.append((rotulo, atual_u, novo_u))
                linha_mudou = True

        if linha_mudou:
            escrita_pendente[linha] = [
                novo_n if novo_n is not None else atual_n,
                novo_u if novo_u is not None else atual_u,
            ]

    # ---- escrita em blocos contiguos (minimiza chamadas COM) --------
    if escrita_pendente:
        linhas = sorted(escrita_pendente)
        blocos, bloco = [], [linhas[0]]
        for ln in linhas[1:]:
            if ln == bloco[-1] + 1:
                bloco.append(ln)
            else:
                blocos.append(bloco)
                bloco = [ln]
        blocos.append(bloco)

        for bloco in blocos:
            matriz = [escrita_pendente[ln] for ln in bloco]
            ini, fim = bloco[0], bloco[-1]
            rng = sh.range(f"{COL_NOVO}{ini}:{COL_USADO}{fim}")
            rng.value = matriz
            rng.api.NumberFormatLocal = FORMATO_PCT
            print(f"  Escritas linhas {ini}:{fim} ({len(bloco)} meses)")

    return {
        "salvar":        sum(contadores.values()) > 0,
        "contadores":    contadores,
        "revisoes_novo": revisoes_novo,
        "revisoes_usado": revisoes_usado,
        "ausentes":      ausentes,
        "ultimo_mes":    max(dados) if dados else None,
    }


# ============================================================
# RELATORIO
# ============================================================
def print_relatorio(r: dict):
    print("\n== Atualizacao IPCA Veiculos — CarRental.xlsm ==")
    if r["ultimo_mes"]:
        ano, mes = r["ultimo_mes"]
        print(f"Ultimo mes disponivel no IBGE: {mes:02d}/{ano}")

    def listar(revisoes, limite=12):
        # 'antigo' e o valor cru que ja estava na celula (pode estar em unidade
        # antiga durante uma migracao); 'novo' e sempre a fracao a ser gravada.
        for mes, antigo, novo in revisoes[:limite]:
            print(f"    - {mes}: {antigo} -> {novo * 100:+.2f}%")
        if len(revisoes) > limite:
            print(f"    ... (+{len(revisoes) - limite} revisoes)")

    c = r["contadores"]
    print(f"\nAutomovel Novo (C):  novas={c['novas_novo']}  revisoes={c['rev_novo']}")
    listar(r["revisoes_novo"])

    print(f"Automovel Usado (D): novas={c['novas_usado']}  revisoes={c['rev_usado']}")
    listar(r["revisoes_usado"])

    if r["ausentes"]:
        print(f"\nMeses retornados fora do range de datas da planilha (ignorados): "
              f"{len(r['ausentes'])}")
        print(f"    (ex.: {', '.join(r['ausentes'][:5])} ...)")

    print(f"\nStatus: {'SUCESSO' if r['salvar'] else 'SEM ALTERACOES'}")


if __name__ == "__main__":
    main()
