"""
Atualização semanal SIDRA 1092 — aba 'Abates' de BeefBR.xlsm

Via xlwings (preserva links externos, formatação e demais abas intactas).
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
XLSX_NAME = "BeefBR.xlsm"
SHEET_NAME = "Abates"

API_URL_TEMPLATE = (
    "https://apisidra.ibge.gov.br/values/t/1092/n1/1/v/284,285"
    "/p/{periods}/c12716/115233,115234,115235"
    "/c18/55,56,111734,111735,57/c12529/118225"
)

CATEGORIA_COL = {  # c18 -> coluna no bloco heads (C:G) e peso (O:S)
    "55":     {"heads": "C", "peso": "O", "nome": "Bois"},
    "56":     {"heads": "D", "peso": "P", "nome": "Vacas"},
    "111734": {"heads": "E", "peso": "Q", "nome": "Novilhos"},
    "111735": {"heads": "F", "peso": "R", "nome": "Novilhas"},
    "57":     {"heads": "G", "peso": "S", "nome": "Vitelos e Vitelas"},
}

MES_NO_TRIMESTRE = {"115233": 1, "115234": 2, "115235": 3}
TIMEOUT = 30
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
            f"Foram encontradas várias planilhas chamadas {XLSX_NAME!r}:\n"
            f"{caminhos}"
        )

    return encontrados[0]


def row_of(ano: int, mes: int) -> int:
    """Linha 3 = jan/97."""
    return 3 + (ano - 1997) * 12 + (mes - 1)


def trimestre_to_mes(trimestre: int, mes_no_tri: int) -> int:
    return (trimestre - 1) * 3 + mes_no_tri


def parse_valor(raw: str, categoria_cod: str, variavel: str):
    """Retorna número ou levanta erro. '...' e '-' viram 0. 'X' aborta."""
    if raw is None:
        return None
    s = str(raw).strip()
    if s in ("...", "-", ""):
        return 0
    if s == "X":
        raise RuntimeError(f"Valor 'X' retornado (categoria {categoria_cod}, var {variavel}).")
    if variavel == "284":
        return int(float(s))
    return float(s)


def parse_valor_safe(raw, cat, var):
    try:
        return parse_valor(raw, cat, var)
    except Exception:
        return None


def fetch_sidra(periods: list[str]) -> list[dict]:
    url = API_URL_TEMPLATE.format(periods=",".join(periods))
    last_err = None
    for tentativa in range(1, MAX_RETRIES + 1):
        try:
            r = requests.get(url, timeout=TIMEOUT)
            r.raise_for_status()
            data = r.json()
            if not isinstance(data, list) or len(data) < 2:
                raise RuntimeError(f"Payload inesperado: {data!r}")
            return data[1:]  # linha 0 é header descritivo
        except Exception as e:
            last_err = e
            if tentativa < MAX_RETRIES:
                time.sleep(RETRY_SLEEP)
    raise RuntimeError(f"API SIDRA falhou em {MAX_RETRIES} tentativas: {last_err}")


def descobre_ultimo_trimestre_disponivel() -> tuple[int, int]:
    """Testa trimestres recentes indo para trás até achar dado."""
    hoje = dt.date.today()
    ano = hoje.year
    tri = (hoje.month - 1) // 3 + 1
    for _ in range(8):
        periodo = f"{ano}{tri:02d}"
        try:
            rows = fetch_sidra([periodo])
            if any(
                parse_valor_safe(r.get("V"), r.get("D1C"), r.get("D2C")) not in (None, 0)
                for r in rows
            ):
                return ano, tri
        except Exception:
            pass
        tri -= 1
        if tri == 0:
            tri = 4
            ano -= 1
    raise RuntimeError("Não foi possível determinar último trimestre disponível.")


# ============================================================
# MAIN
# ============================================================
def main():
    path = localizar_planilha()

    print(f"[{dt.datetime.now():%Y-%m-%d %H:%M}] Abrindo {path.name} via xlwings...")

    app = xw.App(visible=False, add_book=False)
    try:
        # Desabilita atualização de links no nível da aplicação ANTES de abrir
        app.api.AskToUpdateLinks = False
        app.display_alerts = False
        app.screen_updating = False

        wb = app.books.open(
            str(path),
            update_links=0,   # 0 = não atualiza nenhum link (mais explícito que False)
            read_only=False,
        )
        try:
            resultado = processar(wb)
            if resultado["salvar"]:
                # Preserva valores cacheados dos links externos ao salvar
                wb.api.SaveLinkValues = True
                wb.save()
                print("Arquivo salvo.")
            else:
                print("Nada alterado — arquivo NÃO salvo.")
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

    last_row = sh.used_range.last_cell.row
    heads_vals  = sh.range(f"C3:G{last_row}").value
    peso_vals   = sh.range(f"O3:S{last_row}").value
    datas_col_b = sh.range(f"B3:B{last_row}").value

    def diagnostico(matriz):
        primeiro = ultimo = None
        meses_preenchidos = set()
        for i, linha in enumerate(matriz):
            if not isinstance(linha, (list, tuple)):
                linha = [linha]
            if any(v is not None for v in linha):
                if primeiro is None:
                    primeiro = i
                ultimo = i
                meses_preenchidos.add(i)
        gap = (
            primeiro is None
            or primeiro != 0
            or (ultimo is not None and len(meses_preenchidos) != ultimo - primeiro + 1)
        )
        return primeiro, ultimo, gap

    p_h, u_h, gap_h = diagnostico(heads_vals)
    p_p, u_p, gap_p = diagnostico(peso_vals)
    print(f"Heads: primeiro={p_h}, último={u_h}, gap={gap_h}")
    print(f"Peso:  primeiro={p_p}, último={u_p}, gap={gap_p}")

    template_row = 3 + u_h if u_h is not None else None
    print(f"Linha template fixada: {template_row}")

    # ---- período a consultar ----------------------------------------
    ano_ult, tri_ult = descobre_ultimo_trimestre_disponivel()
    print(f"Último trimestre disponível no IBGE: {tri_ult}T{ano_ult % 100:02d}")

    if gap_h or gap_p:
        periods = []
        a, t = 1997, 1
        while (a, t) <= (ano_ult, tri_ult):
            periods.append(f"{a}{t:02d}")
            t += 1
            if t == 5:
                t = 1
                a += 1
    else:
        periods = []
        a, t = ano_ult, tri_ult
        for _ in range(5):
            periods.append(f"{a}{t:02d}")
            t -= 1
            if t == 0:
                t = 4
                a -= 1
        periods.reverse()

    # Evita UnicodeEncodeError em consoles Windows (cp1252) por causa do caractere "→".
    print(f"Consultando {len(periods)} trimestres: {periods[0]} -> {periods[-1]}")

    # ---- buscar e parsear -------------------------------------------
    rows = fetch_sidra(periods)

    dados = {}
    for r in rows:
        var         = r.get("D2C")
        cat         = r.get("D5C")
        mes_tri_cod = r.get("D4C")
        periodo     = r.get("D3C")
        if not all([var, cat, mes_tri_cod, periodo]):
            continue
        ano = int(periodo[:4])
        tri = int(periodo[4:])
        mes = trimestre_to_mes(tri, MES_NO_TRIMESTRE[mes_tri_cod])
        valor = parse_valor(r.get("V"), cat, var)
        key = (ano, mes, cat)
        dados.setdefault(key, {})
        dados[key]["heads" if var == "284" else "peso"] = valor

    # ---- comparar e escrever ----------------------------------------
    contadores = {
        "escritas_novas_heads": 0, "revisoes_heads": 0,
        "escritas_novas_peso":  0, "revisoes_peso":  0,
    }
    revisoes_h = []
    revisoes_p = []
    linhas_novas = set()
    assimetrias = []

    for (ano, mes, cat), vals in sorted(dados.items()):
        linha = row_of(ano, mes)

        idx_b = linha - 3
        if idx_b < 0 or idx_b >= len(datas_col_b):
            raise RuntimeError(f"Linha calculada {linha} fora do intervalo da coluna B.")
        data_b = datas_col_b[idx_b]
        if isinstance(data_b, dt.datetime):
            if data_b.year != ano or data_b.month != mes:
                raise RuntimeError(
                    f"Divergência col B linha {linha}: esperado {mes:02d}/{ano}, "
                    f"encontrado {data_b.month:02d}/{data_b.year}"
                )

        cols = CATEGORIA_COL[cat]
        rotulo_mes = f"{mes:02d}/{ano}"
        categoria_nome = cols["nome"]

        heads_row_idx = linha - 3
        if heads_row_idx < len(heads_vals):
            heads_linha_atual = heads_vals[heads_row_idx]
            if not isinstance(heads_linha_atual, (list, tuple)):
                heads_linha_atual = [heads_linha_atual]
            tinha_heads = any(v is not None for v in heads_linha_atual)
        else:
            tinha_heads = False
        if not tinha_heads:
            linhas_novas.add(linha)

        if "heads" in vals:
            cel = sh.range(f"{cols['heads']}{linha}")
            atual = cel.value
            novo = vals["heads"]
            if atual is None:
                cel.value = novo
                contadores["escritas_novas_heads"] += 1
            elif _diff(atual, novo):
                cel.value = novo
                contadores["revisoes_heads"] += 1
                revisoes_h.append((rotulo_mes, categoria_nome, atual, novo))

        if "peso" in vals:
            cel = sh.range(f"{cols['peso']}{linha}")
            atual = cel.value
            novo = vals["peso"]
            if atual is None:
                cel.value = novo
                contadores["escritas_novas_peso"] += 1
            elif _diff(atual, novo):
                cel.value = novo
                contadores["revisoes_peso"] += 1
                revisoes_p.append((rotulo_mes, categoria_nome, atual, novo))

        if ("heads" in vals) ^ ("peso" in vals):
            assimetrias.append((
                rotulo_mes, categoria_nome,
                "só heads" if "heads" in vals else "só peso",
            ))

    # ---- propagar formato/fórmulas para linhas novas ----------------
    if linhas_novas and template_row is not None:
        tmpl_cm = sh.range(f"C{template_row}:M{template_row}")
        tmpl_ot = sh.range(f"O{template_row}:T{template_row}")

        for linha in sorted(linhas_novas):
            alvo_cm = sh.range(f"C{linha}:M{linha}")
            alvo_ot = sh.range(f"O{linha}:T{linha}")

            tmpl_cm.api.Copy()
            alvo_cm.api.PasteSpecial(Paste=-4122)  # xlPasteFormats
            tmpl_ot.api.Copy()
            alvo_ot.api.PasteSpecial(Paste=-4122)

            sh.range(f"H{template_row}:M{template_row}").api.Copy()
            sh.range(f"H{linha}:M{linha}").api.PasteSpecial(Paste=-4123)  # xlPasteFormulas
            sh.range(f"T{template_row}").api.Copy()
            sh.range(f"T{linha}").api.PasteSpecial(Paste=-4123)

        sh.book.app.api.CutCopyMode = False

    total_mudancas = sum(contadores.values())

    return {
        "salvar": total_mudancas > 0,
        "contadores": contadores,
        "revisoes_h": revisoes_h,
        "revisoes_p": revisoes_p,
        "linhas_novas": sorted(linhas_novas),
        "template_row": template_row,
        "periodos": (periods[0], periods[-1]),
        "ultimo_tri_ibge": f"{tri_ult}T{ano_ult % 100:02d}",
        "assimetrias": assimetrias,
    }


def _diff(a, b) -> bool:
    try:
        return abs(float(a) - float(b)) > 2000
    except (TypeError, ValueError):
        return a != b


# ============================================================
# RELATÓRIO
# ============================================================
def print_relatorio(r: dict):
    print("\n== Atualização SIDRA 1092 — Abates Bovinos ==")
    print(f"Período consultado: {r['periodos'][0]} a {r['periodos'][1]}")
    print(f"Último trimestre disponível no IBGE: {r['ultimo_tri_ibge']}")
    print(f"\nLinhas novas inseridas: {len(r['linhas_novas'])}")
    for linha in r["linhas_novas"]:
        print(f"  - linha {linha}")
    print(f"Linha template usada: {r['template_row']}")

    c = r["contadores"]
    print(f"\nHeads (C:G):")
    print(f"  Escritas novas: {c['escritas_novas_heads']}")
    print(f"  Revisões: {c['revisoes_heads']}")
    for mes, cat, antigo, novo in r["revisoes_h"]:
        print(f"    - {mes} {cat}: {antigo} → {novo}")

    print(f"\nPeso (O:S):")
    print(f"  Escritas novas: {c['escritas_novas_peso']}")
    print(f"  Revisões: {c['revisoes_peso']}")
    for mes, cat, antigo, novo in r["revisoes_p"]:
        print(f"    - {mes} {cat}: {antigo} → {novo}")

    if r["assimetrias"]:
        print(f"\nAssimetrias:")
        for a in r["assimetrias"]:
            print(f"  - {a}")
    else:
        print("\nAssimetrias: nenhuma")

    print(f"\nStatus: {'SUCESSO' if r['salvar'] else 'SEM ALTERAÇÕES'}")


if __name__ == "__main__":
    main()
