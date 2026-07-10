"""
atualizar_beefus.py

Consolida/atualiza a aba principal "BeefUS" (BeefUS.xlsm) para ficar
compatibilizada com as abas auxiliares.

Regras implementadas (conforme solicitado):
1) Abates:
   - Aba auxiliar: "Abates" (dados C:N).
   - Aba principal: "BeefUS" (tabela de abates em D:J).
   - O script compara a última linha com dado em "Abates" (col C) e
     estende a aba "BeefUS" até a mesma referência temporal (offset de linha),
     arrastando fórmulas (e formato) da última linha-template.

2) Preços (na aba "BeefUS"):
   - Coluna M: deve existir fórmula para o mês corrente (linha do mês atual).
     Se estiver vazia, copia a fórmula do mês anterior.
   - Coluna N: é preenchida por outro script (não mexe).
   - Colunas O e P: se estiverem vazias e a coluna N tiver valor na mesma linha,
     copia as fórmulas do mês anterior para O e/ou P.

Requer: pip install xlwings
"""

from __future__ import annotations

import datetime as dt
import os
import sys
from pathlib import Path

import xlwings as xw


# ============================================================
# CONFIG
# ============================================================
DRIVE_ROOT = Path(
    r"G:\.shortcut-targets-by-id\11q1ngdqbySnDMCQSP4fiwPztT6ju2F4Z"
)
XLSX_NAME = "BeefUS.xlsm"

SHEET_MAIN = "BeefUS"
SHEET_ABATES = "Abates"

# BeefUS (principal) começa em linha 5 (ver atualiza_px_bezerro_usa.py).
MAIN_FIRST_DATA_ROW = 5
# Abates (USDA) começa em linha 3 (ver atualiza_abates_usda.py).
ABATES_FIRST_DATA_ROW = 3
# Offset para mapear linha de Abates -> linha equivalente no BeefUS:
# (ex.: jan/2000 em Abates=3, em BeefUS=5 => offset = +2)
ROW_OFFSET_ABATES_TO_MAIN = MAIN_FIRST_DATA_ROW - ABATES_FIRST_DATA_ROW

# Seção Abates no BeefUS
ABATES_MAIN_COL_START = 4   # D
ABATES_MAIN_COL_END = 10    # J

# Preços na BeefUS
COL_DATE_MAIN = 2   # B
COL_M = 13          # M
COL_N = 14          # N
COL_O = 15          # O
COL_P = 16          # P
COL_Q = 17          # Q (EdgeBeef)


# ============================================================
# Excel constants (COM)
# ============================================================
XL_UP = -4162
XL_PASTE_FORMATS = -4122
XL_PASTE_FORMULAS = -4123


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


def col_letter(n: int) -> str:
    """Número de coluna (1-based) → letra Excel."""
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _last_row_with_value(ws: xw.Sheet, col: int, min_row: int) -> int:
    """Última linha (>=min_row) onde a célula em col tem valor não vazio."""
    bottom = ws.api.Cells(ws.api.Rows.Count, col).End(XL_UP).Row
    if bottom < min_row:
        return min_row - 1
    # Se a coluna tiver header ou fórmulas acima, valida descendo um pouco.
    for r in range(bottom, min_row - 1, -1):
        v = ws.api.Cells(r, col).Value
        if v not in (None, ""):
            return r
    return min_row - 1


def _last_row_with_formula(ws: xw.Sheet, col: int, min_row: int) -> int:
    """Última linha (>=min_row) onde a célula em col contém fórmula."""
    bottom = ws.api.Cells(ws.api.Rows.Count, col).End(XL_UP).Row
    if bottom < min_row:
        return min_row - 1
    for r in range(bottom, min_row - 1, -1):
        cell = ws.api.Cells(r, col)
        try:
            if bool(cell.HasFormula):
                return r
        except Exception:
            # alguns tipos de célula podem dar erro em HasFormula
            pass
    return min_row - 1


def _copy_row_formats(ws: xw.Sheet, src_row: int, dst_start: int, dst_end: int, col_start: int, col_end: int) -> None:
    if dst_end < dst_start:
        return
    src = ws.range((src_row, col_start), (src_row, col_end))
    dst = ws.range((dst_start, col_start), (dst_end, col_end))
    src.api.Copy()
    dst.api.PasteSpecial(Paste=XL_PASTE_FORMATS)


def _copy_formulas_for_cols(ws: xw.Sheet, src_row: int, dst_start: int, dst_end: int, col_start: int, col_end: int) -> int:
    """
    Copia (PasteSpecial formulas) apenas das colunas onde o src_row tem fórmula.
    Retorna quantas colunas receberam fórmula (para log).
    """
    if dst_end < dst_start:
        return 0

    formula_cols = []
    for c in range(col_start, col_end + 1):
        cell = ws.api.Cells(src_row, c)
        try:
            if bool(cell.HasFormula):
                formula_cols.append(c)
        except Exception:
            continue

    if not formula_cols:
        return 0

    # Agrupa colunas contíguas para reduzir operações COM.
    groups = []
    g0 = g1 = formula_cols[0]
    for c in formula_cols[1:]:
        if c == g1 + 1:
            g1 = c
        else:
            groups.append((g0, g1))
            g0 = g1 = c
    groups.append((g0, g1))

    for c0, c1 in groups:
        src = ws.range((src_row, c0), (src_row, c1))
        dst = ws.range((dst_start, c0), (dst_end, c1))
        src.api.Copy()
        dst.api.PasteSpecial(Paste=XL_PASTE_FORMULAS)

    return len(formula_cols)


def _find_row_for_year_month(ws: xw.Sheet, date_col: int, first_row: int, year: int, month: int, search_max: int = 2000) -> int | None:
    """
    Encontra a linha onde a coluna de datas (tipicamente B) tem (year, month).
    Retorna None se não encontrar em um intervalo razoável.
    """
    last_used = ws.used_range.last_cell.row
    last_row = min(last_used, first_row + search_max)
    vals = ws.range((first_row, date_col), (last_row, date_col)).value

    # xlwings pode retornar escalar, lista ou lista de listas.
    if vals is None:
        return None
    if not isinstance(vals, list):
        vals = [vals]
    if vals and isinstance(vals[0], list):
        vals = [v[0] if v else None for v in vals]

    for i, v in enumerate(vals):
        if isinstance(v, dt.datetime):
            if v.year == year and v.month == month:
                return first_row + i
        elif isinstance(v, dt.date):
            if v.year == year and v.month == month:
                return first_row + i
    return None


def sincronizar_abates(ws_main: xw.Sheet, ws_abates: xw.Sheet) -> dict:
    """
    Faz BeefUS (D:J) acompanhar o último mês/linha da aba Abates, arrastando fórmulas.
    """
    # 1) Descobre "última linha com dado" em Abates usando col C (steers).
    last_abates = _last_row_with_value(ws_abates, col=3, min_row=ABATES_FIRST_DATA_ROW)
    if last_abates < ABATES_FIRST_DATA_ROW:
        return {"mudou": False, "msg": "Abates: nenhuma linha com dado em col C."}

    target_main = last_abates + ROW_OFFSET_ABATES_TO_MAIN

    # 2) Template = última linha que já tem fórmula na coluna D (na seção Abates).
    template_main = _last_row_with_formula(ws_main, col=ABATES_MAIN_COL_START, min_row=MAIN_FIRST_DATA_ROW)
    if template_main < MAIN_FIRST_DATA_ROW:
        return {"mudou": False, "msg": "BeefUS: nenhuma linha-template com fórmula em D."}

    if target_main <= template_main:
        return {"mudou": False, "msg": f"Abates: já alinhado (última linha BeefUS={template_main})."}

    dst_start = template_main + 1
    dst_end = target_main

    # 3) Copia formatos numa faixa ampla (A:P) para as novas linhas.
    _copy_row_formats(ws_main, src_row=template_main, dst_start=dst_start, dst_end=dst_end, col_start=1, col_end=COL_P)

    # 4) Copia fórmulas (A:P) apenas onde template tem fórmula.
    ncols = _copy_formulas_for_cols(ws_main, src_row=template_main, dst_start=dst_start, dst_end=dst_end, col_start=1, col_end=COL_P)

    # Evita "marching ants"
    try:
        ws_main.book.app.api.CutCopyMode = False
    except Exception:
        pass

    return {
        "mudou": True,
        "novas_linhas": dst_end - template_main,
        "template_row": template_main,
        "target_row": target_main,
        "formula_cols": ncols,
        "msg": f"Abates: +{dst_end - template_main} linha(s) em BeefUS (até {target_main}).",
    }


def ajustar_precos(ws_main: xw.Sheet, today: dt.date) -> dict:
    """
    Ajusta M (mês corrente) e completa O/P quando N tiver valor.
    """
    # Linha do mês corrente pela data em col B
    row_cur = _find_row_for_year_month(ws_main, date_col=COL_DATE_MAIN, first_row=MAIN_FIRST_DATA_ROW, year=today.year, month=today.month)
    if row_cur is None:
        return {"mudou": False, "msg": "Preços: não encontrei a linha do mês atual na col B."}

    # Linha do mês anterior
    prev = (today.replace(day=1) - dt.timedelta(days=1))
    row_prev = _find_row_for_year_month(ws_main, date_col=COL_DATE_MAIN, first_row=MAIN_FIRST_DATA_ROW, year=prev.year, month=prev.month)
    if row_prev is None:
        # fallback: usa linha imediatamente anterior
        row_prev = row_cur - 1
        if row_prev < MAIN_FIRST_DATA_ROW:
            return {"mudou": False, "msg": "Preços: sem linha anterior para usar como template."}

    mudou = False
    acoes = []

    # --- Coluna M: preenche do último mês com fórmula até o mês corrente ---
    last_m_row = _last_row_with_formula(ws_main, col=COL_M, min_row=MAIN_FIRST_DATA_ROW)
    if last_m_row >= MAIN_FIRST_DATA_ROW and last_m_row < row_cur:
        dst_start_m = last_m_row + 1
        src = ws_main.range((last_m_row, COL_M), (last_m_row, COL_M))
        dst = ws_main.range((dst_start_m, COL_M), (row_cur, COL_M))
        src.api.Copy()
        dst.api.PasteSpecial(Paste=XL_PASTE_FORMULAS)
        src.api.Copy()
        dst.api.PasteSpecial(Paste=XL_PASTE_FORMATS)
        mudou = True
        acoes.append(f"M: fórmula copiada linhas {dst_start_m}->{row_cur} ({row_cur - last_m_row} mes(es))")

    # --- Coluna Q (EdgeBeef): idêntico ao M — até o mês corrente ---
    last_q_row = _last_row_with_formula(ws_main, col=COL_Q, min_row=MAIN_FIRST_DATA_ROW)
    if last_q_row >= MAIN_FIRST_DATA_ROW and last_q_row < row_cur:
        dst_start_q = last_q_row + 1
        src = ws_main.range((last_q_row, COL_Q), (last_q_row, COL_Q))
        dst = ws_main.range((dst_start_q, COL_Q), (row_cur, COL_Q))
        src.api.Copy()
        dst.api.PasteSpecial(Paste=XL_PASTE_FORMULAS)
        src.api.Copy()
        dst.api.PasteSpecial(Paste=XL_PASTE_FORMATS)
        mudou = True
        acoes.append(f"Q: fórmula copiada linhas {dst_start_q}->{row_cur} ({row_cur - last_q_row} mes(es))")

    # --- Colunas O/P: se N tem valor e O/P vazias, puxa fórmula ---
    n_val = ws_main.api.Cells(row_cur, COL_N).Value
    if n_val not in (None, ""):
        for col, label in [(COL_O, "O"), (COL_P, "P")]:
            cell = ws_main.api.Cells(row_cur, col)
            has_formula = False
            try:
                has_formula = bool(cell.HasFormula)
            except Exception:
                has_formula = False
            val = cell.Value
            if (not has_formula) and (val in (None, "")):
                src = ws_main.range((row_prev, col), (row_prev, col))
                dst = ws_main.range((row_cur, col), (row_cur, col))
                src.api.Copy()
                dst.api.PasteSpecial(Paste=XL_PASTE_FORMULAS)
                src.api.Copy()
                dst.api.PasteSpecial(Paste=XL_PASTE_FORMATS)
                mudou = True
                acoes.append(f"{label}: fórmula copiada {row_prev}->{row_cur} (N preenchida)")

    try:
        ws_main.book.app.api.CutCopyMode = False
    except Exception:
        pass

    if not mudou:
        return {"mudou": False, "msg": "Preços: nada para ajustar."}
    return {"mudou": True, "msg": "Preços: " + "; ".join(acoes), "acoes": acoes, "row_cur": row_cur, "row_prev": row_prev}


def main() -> None:
    try:
        path = localizar_planilha()
    except (FileNotFoundError, RuntimeError) as exc:
        print(f"ERRO: {exc}")
        raise SystemExit(1)

    today = dt.date.today()
    print(f"[{dt.datetime.now():%Y-%m-%d %H:%M}] Atualizando BeefUS (consolidação) — {today:%d/%m/%Y}")
    print(f"Arquivo: {path.name}")

    app = xw.App(visible=False, add_book=False)
    try:
        app.api.AskToUpdateLinks = False
        app.display_alerts = False
        app.screen_updating = False

        wb = app.books.open(str(path), update_links=0, read_only=False)
        try:
            ws_main = wb.sheets[SHEET_MAIN]
            ws_abates = wb.sheets[SHEET_ABATES]

            r1 = sincronizar_abates(ws_main, ws_abates)
            print(" - " + r1["msg"])

            r2 = ajustar_precos(ws_main, today=today)
            print(" - " + r2["msg"])

            if bool(r1.get("mudou")) or bool(r2.get("mudou")):
                wb.api.SaveLinkValues = True
                wb.save()
                print("Salvo.")
            else:
                print("Nada alterado — NÃO salvou.")
        finally:
            wb.close()
    finally:
        app.quit()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        raise SystemExit(130)
