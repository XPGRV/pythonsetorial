"""
Atualização trimestral SIDRA 1094 — aba 'Production' de FrangoBR.xlsm

Escreve nas colunas:
  C = Animais abatidos (cabeças)   — variável 284
  D = Peso total das carcaças (kg) — variável 285

Via xlwings (preserva links externos, formatação e demais abas intactas).
"""

import sys
import time
import datetime as dt
from copy import copy
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import drive_io
import ibge_sidra


# ============================================================
# CONFIG
# ============================================================
DRIVE_ROOT = Path(
    r"G:\.shortcut-targets-by-id\11q1ngdqbySnDMCQSP4fiwPztT6ju2F4Z"
)
XLSX_NAME  = "FrangoBR.xlsm"
# Le pela API servicodados do IBGE (ibge_sidra.py): o apisidra.ibge.gov.br
# passou a bloquear clientes que nao sao navegador (Cloudflare, set/2026).
CONSULTA = dict(
    tabela="1094",
    variaveis=["284", "285"],
    classificacoes=[("12716", ["115233", "115234", "115235"]), ("12529", ["118225"])],
)
SHEET_NAME = "Production"


COL_HEADS  = "C"   # Animais abatidos (cabeças)
COL_PESO   = "D"   # Peso total das carcaças (kg)
COL_DATAS  = "B"   # Datas mensais — usada para validação

ANO_BASE = 2000
ROW_BASE = 3       # linha 3 = jan/2000

MES_NO_TRIMESTRE = {"115233": 1, "115234": 2, "115235": 3}
TIMEOUT     = 30
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
            f"Planilha {XLSX_NAME!r} nao encontrada dentro de {DRIVE_ROOT}"
        )

    if len(encontrados) > 1:
        caminhos = "\n".join(f"  - {path}" for path in encontrados)
        raise RuntimeError(
            f"Foram encontradas varias planilhas chamadas {XLSX_NAME!r}:\n"
            f"{caminhos}"
        )

    return encontrados[0]


def row_of(ano: int, mes: int) -> int:
    """Linha 3 = jan/2000."""
    return ROW_BASE + (ano - ANO_BASE) * 12 + (mes - 1)


def trimestre_to_mes(trimestre: int, mes_no_tri: int) -> int:
    return (trimestre - 1) * 3 + mes_no_tri


def parse_valor(raw, variavel: str):
    if raw is None:
        return None
    s = str(raw).strip()
    if s in ("...", "-", ""):
        return 0
    if s == "X":
        raise RuntimeError(f"Valor 'X' retornado pela API (var {variavel}).")
    if variavel == "284":
        return int(float(s))
    return float(s)


def parse_valor_safe(raw, var):
    try:
        return parse_valor(raw, var)
    except Exception:
        return None


def fetch_sidra(periods: list[str]) -> list[dict]:
    last_err = None
    for tentativa in range(1, MAX_RETRIES + 1):
        try:
            data = ibge_sidra.consultar(periodos=periods, timeout=TIMEOUT, **CONSULTA)
            if not data:
                raise RuntimeError(f"Sem dado publicado para {periods}")
            return data
        except Exception as e:
            last_err = e
            if tentativa < MAX_RETRIES:
                time.sleep(RETRY_SLEEP)
    raise RuntimeError(f"API SIDRA falhou em {MAX_RETRIES} tentativas: {last_err}")


def descobre_ultimo_trimestre_disponivel() -> tuple[int, int]:
    """Testa trimestres recentes indo para trás até achar dado."""
    hoje = dt.date.today()
    ano  = hoje.year
    tri  = (hoje.month - 1) // 3 + 1
    for _ in range(8):
        periodo = f"{ano}{tri:02d}"
        try:
            rows = fetch_sidra([periodo])
            if any(
                parse_valor_safe(r.get("V"), r.get("D2C")) not in (None, 0)
                for r in rows
            ):
                return ano, tri
        except Exception:
            pass
        tri -= 1
        if tri == 0:
            tri = 4
            ano -= 1
    raise RuntimeError("Nao foi possivel determinar ultimo trimestre disponivel.")


def _diff(a, b) -> bool:
    try:
        return abs(float(a) - float(b)) > 5000
    except (TypeError, ValueError):
        return a != b


# ============================================================
# MAIN
# ============================================================
def main():
    fid = drive_io.find_file_id(XLSX_NAME)
    path = drive_io.download_to_temp(fid)

    print(f"[{dt.datetime.now():%Y-%m-%d %H:%M}] Abrindo FrangoBR.xlsm via openpyxl...")
    wb = openpyxl.load_workbook(path, keep_vba=True, data_only=False)
    try:
        resultado = processar(wb)
        if resultado["salvar"]:
            wb.save(path)
            drive_io.update_file(fid, path)   # sobe de volta pro Drive
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
    ws = wb[SHEET_NAME]
    cB = column_index_from_string(COL_DATAS)
    cC = column_index_from_string(COL_HEADS)
    cD = column_index_from_string(COL_PESO)

    last_row   = ws.max_row
    heads_vals = [ws.cell(r, cC).value for r in range(ROW_BASE, last_row + 1)]
    peso_vals  = [ws.cell(r, cD).value for r in range(ROW_BASE, last_row + 1)]
    datas_col  = [ws.cell(r, cB).value for r in range(ROW_BASE, last_row + 1)]

    def diagnostico(lista):
        primeiro = ultimo = None
        preenchidas = set()
        for i, v in enumerate(lista):
            if v is not None:
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

    p_h, u_h, gap_h = diagnostico(heads_vals)
    p_p, u_p, gap_p = diagnostico(peso_vals)
    print(f"Heads (C): primeiro={p_h}, ultimo={u_h}, gap={gap_h}")
    print(f"Peso  (D): primeiro={p_p}, ultimo={u_p}, gap={gap_p}")

    template_row = (ROW_BASE + u_h) if u_h is not None else None
    print(f"Linha template: {template_row}")

    # ---- período a consultar ----------------------------------------
    ano_ult, tri_ult = descobre_ultimo_trimestre_disponivel()
    print(f"Ultimo trimestre disponivel no IBGE: {tri_ult}T{ano_ult % 100:02d}")

    if gap_h or gap_p:
        # recarga completa desde jan/2000
        periods = []
        a, t = ANO_BASE, 1
        while (a, t) <= (ano_ult, tri_ult):
            periods.append(f"{a}{t:02d}")
            t += 1
            if t == 5:
                t = 1
                a += 1
    else:
        # apenas os últimos 5 trimestres
        periods = []
        a, t = ano_ult, tri_ult
        for _ in range(5):
            periods.append(f"{a}{t:02d}")
            t -= 1
            if t == 0:
                t = 4
                a -= 1
        periods.reverse()

    print(f"Consultando {len(periods)} trimestres: {periods[0]} -> {periods[-1]}")

    # ---- buscar e parsear -------------------------------------------
    rows = fetch_sidra(periods)

    dados = {}
    for r in rows:
        var         = r.get("D2C")
        mes_tri_cod = r.get("D4C")
        periodo     = r.get("D3C")
        if not all([var, mes_tri_cod, periodo]):
            continue
        if mes_tri_cod not in MES_NO_TRIMESTRE:
            continue  # pula "Total do trimestre" (115236)
        ano = int(periodo[:4])
        tri = int(periodo[4:])
        mes = trimestre_to_mes(tri, MES_NO_TRIMESTRE[mes_tri_cod])
        if ano < ANO_BASE:
            continue
        valor = parse_valor(r.get("V"), var)
        key   = (ano, mes)
        dados.setdefault(key, {})
        dados[key]["heads" if var == "284" else "peso"] = valor

    # ---- comparar e escrever ----------------------------------------
    contadores = {
        "escritas_novas_heads": 0, "revisoes_heads": 0,
        "escritas_novas_peso":  0, "revisoes_peso":  0,
    }
    revisoes_h  = []
    revisoes_p  = []
    linhas_novas = set()
    escrita_pendente: dict[int, list] = {}  # linha -> [heads, peso]

    for (ano, mes), vals in sorted(dados.items()):
        linha = row_of(ano, mes)
        idx   = linha - ROW_BASE

        # validação da coluna de datas
        if 0 <= idx < len(datas_col):
            data_b = datas_col[idx]
            if isinstance(data_b, dt.datetime):
                if data_b.year != ano or data_b.month != mes:
                    raise RuntimeError(
                        f"Divergencia col {COL_DATAS} linha {linha}: "
                        f"esperado {mes:02d}/{ano}, "
                        f"encontrado {data_b.month:02d}/{data_b.year}"
                    )

        rotulo = f"{mes:02d}/{ano}"

        atual_h = heads_vals[idx] if idx < len(heads_vals) else None
        atual_p = peso_vals[idx]  if idx < len(peso_vals)  else None

        if atual_h is None and atual_p is None:
            linhas_novas.add(linha)

        novo_h = vals.get("heads")
        novo_p = vals.get("peso")

        linha_mudou = False

        if novo_h is not None:
            if atual_h is None:
                contadores["escritas_novas_heads"] += 1
                linha_mudou = True
            elif _diff(atual_h, novo_h):
                contadores["revisoes_heads"] += 1
                revisoes_h.append((rotulo, atual_h, novo_h))
                linha_mudou = True

        if novo_p is not None:
            if atual_p is None:
                contadores["escritas_novas_peso"] += 1
                linha_mudou = True
            elif _diff(atual_p, novo_p):
                contadores["revisoes_peso"] += 1
                revisoes_p.append((rotulo, atual_p, novo_p))
                linha_mudou = True

        if linha_mudou:
            escrita_pendente[linha] = [
                novo_h if novo_h is not None else atual_h,
                novo_p if novo_p is not None else atual_p,
            ]

    # escrita em blocos contíguos (minimiza chamadas COM)
    if escrita_pendente:
        linhas_ord  = sorted(escrita_pendente)
        blocos: list[list[int]] = []
        bloco_atual = [linhas_ord[0]]
        for ln in linhas_ord[1:]:
            if ln == bloco_atual[-1] + 1:
                bloco_atual.append(ln)
            else:
                blocos.append(bloco_atual)
                bloco_atual = [ln]
        blocos.append(bloco_atual)

        for bloco in blocos:
            ini, fim = bloco[0], bloco[-1]
            for ln in bloco:
                h_val, p_val = escrita_pendente[ln]
                ws.cell(ln, cC).value = h_val
                ws.cell(ln, cD).value = p_val
            print(f"  Escritas linhas {ini}:{fim} ({len(bloco)} meses)")

    # propagar formato de C:D para linhas novas
    if linhas_novas and template_row is not None:
        novas_ord = sorted(linhas_novas)
        blocos_fmt: list[tuple[int, int]] = []
        ini = fim = novas_ord[0]
        for ln in novas_ord[1:]:
            if ln == fim + 1:
                fim = ln
            else:
                blocos_fmt.append((ini, fim))
                ini = fim = ln
        blocos_fmt.append((ini, fim))

        for ini, fim in blocos_fmt:
            for ln in range(ini, fim + 1):
                for c in (cC, cD):
                    src = ws.cell(template_row, c)
                    dst = ws.cell(ln, c)
                    dst.number_format = src.number_format
                    dst.font = copy(src.font)
                    dst.fill = copy(src.fill)
                    dst.border = copy(src.border)
                    dst.alignment = copy(src.alignment)

    return {
        "salvar":        sum(contadores.values()) > 0,
        "contadores":    contadores,
        "revisoes_h":    revisoes_h,
        "revisoes_p":    revisoes_p,
        "linhas_novas":  sorted(linhas_novas),
        "template_row":  template_row,
        "periodos":      (periods[0], periods[-1]),
        "ultimo_tri_ibge": f"{tri_ult}T{ano_ult % 100:02d}",
    }


# ============================================================
# RELATÓRIO
# ============================================================
def print_relatorio(r: dict):
    print("\n== Atualizacao SIDRA 1094 — Abates de Frango ==")
    print(f"Periodo consultado: {r['periodos'][0]} a {r['periodos'][1]}")
    print(f"Ultimo trimestre disponivel no IBGE: {r['ultimo_tri_ibge']}")
    print(f"\nLinhas novas inseridas: {len(r['linhas_novas'])}")
    for linha in r["linhas_novas"]:
        print(f"  - linha {linha}")
    print(f"Linha template usada: {r['template_row']}")

    c = r["contadores"]
    print(f"\nHeads (C):")
    print(f"  Escritas novas: {c['escritas_novas_heads']}")
    print(f"  Revisoes: {c['revisoes_heads']}")
    for mes, antigo, novo in r["revisoes_h"]:
        print(f"    - {mes}: {antigo:,.0f} -> {novo:,.0f}")

    print(f"\nPeso (D):")
    print(f"  Escritas novas: {c['escritas_novas_peso']}")
    print(f"  Revisoes: {c['revisoes_peso']}")
    for mes, antigo, novo in r["revisoes_p"]:
        print(f"    - {mes}: {antigo:,.0f} -> {novo:,.0f}")

    print(f"\nStatus: {'SUCESSO' if r['salvar'] else 'SEM ALTERACOES'}")


if __name__ == "__main__":
    main()
