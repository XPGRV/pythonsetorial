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

Escrita via openpyxl + drive_io, para rodar no runner do GitHub (que nao tem
Excel). O motivo original de usar xlwings -- as formulas de array do Bloomberg na
aba 'Peers' -- e coberto hoje pelo drive_io._fix_single_cell_array_formulas, que
roda no upload e evita o "Repaired Records" do Excel. Testado: a aba Peers sai do
round-trip com as 6 BDH intactas, o VBA preservado e as duas abas no lugar.

ATENCAO -- esta planilha DEPENDE do passe do Excel logo depois da coleta. Das
cinco colunas que o dashboard le, quatro sao formulas encadeadas:
    F = E - ajuste          (usado %MoM)
    J = J[-1] * (1 + D)     (indice do novo, acumulado)
    L = L[-1] * (1 + F)     (indice do usado, acumulado)
    M = L / J - 1           (spread)
O openpyxl preserva a formula mas descarta o valor em cache, e o parser do
dashboard le justamente o cache. Coletar sem recalcular no Excel derruba o
dataset 'rental' para 0/24 pontos -- verificado. Na esteira isso e garantido
pelo passo 3; rodando o workflow individual, o dado so volta no ciclo seguinte.
"""

import sys
import time
import datetime as dt
from pathlib import Path

import requests
import openpyxl
from openpyxl.utils import column_index_from_string

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import drive_io


# ============================================================
# CONFIG
# ============================================================
XLSX_NAME  = "CarRental.xlsm"
SHEET_NAME = "Preço Carros"

# Corrigido em 18/08/2026: apontavam para C e D, mas C e uma coluna
# separadora (vazia nas 321 linhas de dado) e D ja contem o Novo. Gravar no
# mapeamento antigo poria Usado por cima da serie de Novo, desde jan/2000.
# O cabecalho da planilha e o parser do dashboard confirmam D e E.
COL_NOVO  = "D"   # Automovel Novo  — %MoM
COL_USADO = "E"   # Automovel Usado — %MoM
COL_DATA  = "B"   # datas mensais — usada para validacao

ANO_BASE = 2000
ROW_BASE = 5      # linha 5 = jan/2000

VARIAVEL   = "63"       # variacao mensal (%)
# Formato aplicado as celulas: fracao exibida como percentual com 2 casas.
# Usa-se NumberFormatLocal com a string pt-BR (virgula decimal) porque este
# Excel roda em locale 416; passar "0.00%" faria o Excel ler o "." como
# separador de milhar e gravar "#.000%" (sem casas decimais).
# Invariante, nao local: o openpyxl grava o formato cru no XML, ao contrario
FORMATO_PCT = "0.00%"

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
def valor_existente(v):
    """Valor ja na planilha como float, ou None se nao for numero.

    As colunas D e E usam '-' como marcador de mes sem dado (23 celulas, nos
    primeiros anos da serie). Antes isto nao aparecia porque o script lia a
    coluna C, que e vazia; com o mapeamento corrigido o '-' chega ate aqui e o
    float() estourava.
    """
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    texto = str(v).strip()
    if texto in ("", "-", "--"):
        return None
    try:
        return float(texto.replace(",", "."))
    except ValueError:
        return None


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
    # Sob o cloud_pipeline isto reaproveita o download unico da planilha; rodando
    # sozinho, baixa e sobe de verdade. Nos dois casos, sem Excel instalado.
    file_id = drive_io.find_file_id(XLSX_NAME)
    path = drive_io.download_to_temp(file_id)
    print(f"[{dt.datetime.now():%Y-%m-%d %H:%M}] Abrindo {XLSX_NAME}...")

    wb = openpyxl.load_workbook(path, keep_vba=True, data_only=False)
    try:
        resultado = processar(wb)
        if resultado["salvar"]:
            wb.save(path)
            drive_io.update_file(file_id, path)
            print("Arquivo salvo.")
        else:
            print("Nada alterado — arquivo NAO salvo.")
    finally:
        wb.close()

    print_relatorio(resultado)


# ============================================================
# PROCESSAMENTO
# ============================================================
def processar(wb) -> dict:
    sh = wb[SHEET_NAME]

    # Equivale ao .end("down") do xlwings: desce pela coluna de datas ate a
    # primeira lacuna.
    last_row = ROW_BASE
    while sh[f"{COL_DATA}{last_row + 1}"].value is not None:
        last_row += 1

    def coluna(letra):
        return [
            sh[f"{letra}{row}"].value for row in range(ROW_BASE, last_row + 1)
        ]

    novo_vals  = coluna(COL_NOVO)
    usado_vals = coluna(COL_USADO)
    datas_col  = coluna(COL_DATA)

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

        # bruto_* preserva o conteudo original da celula (inclusive o '-'),
        # atual_* e a versao numerica usada so para comparar.
        bruto_n = novo_vals[idx]  if idx < len(novo_vals)  else None
        bruto_u = usado_vals[idx] if idx < len(usado_vals) else None
        atual_n = valor_existente(bruto_n)
        atual_u = valor_existente(bruto_u)
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
            # Sem dado novo, devolve a celula como estava -- nao troca '-'
            # por vazio so porque o vizinho mudou.
            escrita_pendente[linha] = [
                novo_n if novo_n is not None else bruto_n,
                novo_u if novo_u is not None else bruto_u,
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
            ini, fim = bloco[0], bloco[-1]
            base = column_index_from_string(COL_NOVO)
            for ln in bloco:
                for deslocamento, valor in enumerate(escrita_pendente[ln]):
                    celula = sh.cell(row=ln, column=base + deslocamento)
                    celula.value = valor
                    celula.number_format = FORMATO_PCT
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
    print(f"\nAutomovel Novo ({COL_NOVO}):  novas={c['novas_novo']}  revisoes={c['rev_novo']}")
    listar(r["revisoes_novo"])

    print(f"Automovel Usado ({COL_USADO}): novas={c['novas_usado']}  revisoes={c['rev_usado']}")
    listar(r["revisoes_usado"])

    if r["ausentes"]:
        print(f"\nMeses retornados fora do range de datas da planilha (ignorados): "
              f"{len(r['ausentes'])}")
        print(f"    (ex.: {', '.join(r['ausentes'][:5])} ...)")

    print(f"\nStatus: {'SUCESSO' if r['salvar'] else 'SEM ALTERACOES'}")


if __name__ == "__main__":
    main()
