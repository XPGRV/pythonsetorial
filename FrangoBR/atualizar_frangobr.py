"""
atualizar_frangobr.py

FrangoBR — atualiza formulas em duas abas:

Aba BBG_Dados:
  Colunas J, K, L alinhadas com a ultima linha com dado em cols E->I.

Aba FrangoBR:
  (i)   Mes corrente:  D->I, L, O->R, T  — ate o mes corrente
  (ii)  SIF:          W                  — 2 meses antes do ultimo dado em Production col F
  (iii) IBGE:         X, Y               — ate o ultimo dado em Production cols C, D
  (iv)  APINCO:       Z                  — ate o ultimo dado em Production col J

Mapeamento de linhas:
  FrangoBR:  linha 5 = jan/2000  ->  row = 5 + (ano-2000)*12 + (mes-1)
  Production: linha 3 = jan/2000  ->  FrangoBR_row = Production_row + 2

Requer: pip install pywin32
"""

import sys
import datetime
from pathlib import Path

DRIVE_ROOT = Path(
    r"G:\.shortcut-targets-by-id\11q1ngdqbySnDMCQSP4fiwPztT6ju2F4Z"
)
XLSX_NAME = "FrangoBR.xlsm"

XL_UP          = -4162
XL_PASTE_ALL   = -4104
XL_CALC_MANUAL = -4135
XL_CALC_AUTO   = -4105


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


def col_letter(n: int) -> str:
    result = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def frangobr_row(year: int, month: int) -> int:
    """Linha no FrangoBR para ano/mes. Linha 5 = jan/2000."""
    return 5 + (year - 2000) * 12 + (month - 1)


def production_row_to_ym(row: int):
    """Linha da aba Production -> (ano, mes). Linha 3 = jan/2000."""
    offset = row - 3
    return 2000 + offset // 12, offset % 12 + 1


def prev_month(year: int, month: int):
    return (year - 1, 12) if month == 1 else (year, month - 1)


def format_ym(year: int, month: int) -> str:
    return f"{month:02d}/{year}"


def last_formula_row(ws, col: int, start: int = 5) -> int:
    """Ultima linha em 'col' com formula, buscando de baixo para cima."""
    last = ws.Cells(ws.Rows.Count, col).End(XL_UP).Row
    if last < start:
        return start - 1
    for r in range(last, start - 1, -1):
        if ws.Cells(r, col).HasFormula:
            return r
    return start - 1


def extend_section(ws, src_row: int, col_ranges: list, dst_end: int, excel) -> int:
    if dst_end <= src_row:
        return 0
    dst_start = src_row + 1
    for c_start, c_end in col_ranges:
        rng_src = f"{col_letter(c_start)}{src_row}:{col_letter(c_end)}{src_row}"
        rng_dst = f"{col_letter(c_start)}{dst_start}:{col_letter(c_end)}{dst_end}"
        ws.Range(rng_src).Copy()
        ws.Range(rng_dst).PasteSpecial(Paste=XL_PASTE_ALL)
        excel.CutCopyMode = False
    return dst_end - src_row


# ── BBG_Dados ─────────────────────────────────────────────────────────────────

def atualizar_bbg_dados(ws_bbg, excel) -> None:
    """Colunas J, K, L alinhadas com a ultima linha com dado em cols E->I (ref: col E)."""
    last_data = ws_bbg.Cells(ws_bbg.Rows.Count, 5).End(XL_UP).Row
    last_jkl  = last_formula_row(ws_bbg, col=10, start=2)

    if last_jkl < 2:
        print("  BBG_Dados J->L: nenhuma linha template em col J — pulando.")
        return

    if last_data > last_jkl:
        n = extend_section(ws_bbg, last_jkl, [(10, 12)], last_data, excel)
        print(f"  BBG_Dados J->L: +{n} linha(s) -> linha {last_data}")
    else:
        print(f"  BBG_Dados J->L: ja atualizado (linha {last_jkl})")


# ── FrangoBR ──────────────────────────────────────────────────────────────────

def atualizar_mes_corrente(ws_frango, today: datetime.date, excel) -> None:
    """(i) Cols D->I, L, O->R, T — ate o mes corrente."""
    target = frangobr_row(today.year, today.month)
    last   = last_formula_row(ws_frango, col=4)  # col D

    if last < 5:
        print("  (i)  Mes corrente: nenhuma linha template em col D — pulando.")
        return

    if target > last:
        n = extend_section(
            ws_frango, last,
            [(4, 9), (12, 12), (15, 18), (20, 20)],  # D->I | L | O->R | T
            target, excel
        )
        print(f"  (i)  Mes corrente: +{n} mes(es) -> {format_ym(today.year, today.month)}")
    else:
        print(f"  (i)  Mes corrente: ja atualizado (linha {last})")


def atualizar_sif(ws_frango, ws_prod, excel) -> None:
    """(ii) Col W — 2 meses antes do ultimo dado SIF em Production col F."""
    bottom = ws_prod.Cells(ws_prod.Rows.Count, 6).End(XL_UP).Row
    last_prod = 0
    for r in range(bottom, 2, -1):
        v = ws_prod.Cells(r, 6).Value
        if isinstance(v, (int, float)) and v > 0:
            last_prod = r
            break

    if last_prod == 0:
        print("  (ii) SIF: sem dados em Production col F — pulando.")
        return

    yr, mo = production_row_to_ym(last_prod)
    yr, mo = prev_month(yr, mo)
    yr, mo = prev_month(yr, mo)
    target = frangobr_row(yr, mo)
    last   = last_formula_row(ws_frango, col=23)  # col W

    if last < 5:
        print("  (ii) SIF: nenhuma linha template em col W — pulando.")
        return

    if target > last:
        n = extend_section(ws_frango, last, [(23, 23)], target, excel)
        print(f"  (ii) SIF: +{n} mes(es) -> {format_ym(yr, mo)}")
    else:
        print(f"  (ii) SIF: ja atualizado (linha {last})")


def atualizar_ibge(ws_frango, ws_prod, excel) -> None:
    """(iii) Cols X, Y — ate o ultimo dado IBGE em Production col C."""
    bottom = ws_prod.Cells(ws_prod.Rows.Count, 3).End(XL_UP).Row
    last_prod = 0
    for r in range(bottom, 2, -1):
        v = ws_prod.Cells(r, 3).Value
        if isinstance(v, (int, float)) and v > 0:
            last_prod = r
            break

    if last_prod == 0:
        print("  (iii) IBGE: sem dados em Production col C — pulando.")
        return

    target = last_prod + 2  # Production linha 3 = FrangoBR linha 5
    last   = last_formula_row(ws_frango, col=24)  # col X

    if last < 5:
        print("  (iii) IBGE: nenhuma linha template em col X — pulando.")
        return

    if target > last:
        yr, mo = production_row_to_ym(last_prod)
        n = extend_section(ws_frango, last, [(24, 25)], target, excel)
        print(f"  (iii) IBGE: +{n} mes(es) -> {format_ym(yr, mo)}")
    else:
        print(f"  (iii) IBGE: ja atualizado (linha {last})")


def atualizar_apinco(ws_frango, ws_prod, excel) -> None:
    """(iv) Col Z — ate o ultimo dado APINCO em Production col J."""
    bottom = ws_prod.Cells(ws_prod.Rows.Count, 10).End(XL_UP).Row
    last_prod = 0
    for r in range(bottom, 2, -1):
        v = ws_prod.Cells(r, 10).Value
        if isinstance(v, (int, float)) and v > 0:
            last_prod = r
            break

    if last_prod == 0:
        print("  (iv) APINCO: sem dados em Production col J — pulando.")
        return

    target = last_prod + 2
    last   = last_formula_row(ws_frango, col=26)  # col Z

    if last < 5:
        print("  (iv) APINCO: nenhuma linha template em col Z — pulando.")
        return

    if target > last:
        yr, mo = production_row_to_ym(last_prod)
        n = extend_section(ws_frango, last, [(26, 26)], target, excel)
        print(f"  (iv) APINCO: +{n} mes(es) -> {format_ym(yr, mo)}")
    else:
        print(f"  (iv) APINCO: ja atualizado (linha {last})")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    try:
        import win32com.client
    except ImportError:
        print("ERRO: pip install pywin32")
        sys.exit(1)

    try:
        path = localizar_planilha()
    except (FileNotFoundError, RuntimeError) as exc:
        print(f"ERRO: {exc}")
        sys.exit(1)

    today = datetime.date.today()
    print(f"FrangoBR — {today:%d/%m/%Y}\n")

    excel = win32com.client.DispatchEx("Excel.Application")
    try: excel.Visible        = False
    except Exception: pass
    try: excel.DisplayAlerts  = False
    except Exception: pass
    try: excel.ScreenUpdating = False
    except Exception: pass

    try:
        wb        = excel.Workbooks.Open(str(path))
        excel.Calculation = XL_CALC_MANUAL
        ws_bbg    = wb.Sheets("BBG_Dados")
        ws_frango = wb.Sheets("FrangoBR")
        ws_prod   = wb.Sheets("Production")

        atualizar_bbg_dados(ws_bbg, excel)
        atualizar_mes_corrente(ws_frango, today, excel)
        atualizar_sif(ws_frango, ws_prod, excel)
        atualizar_ibge(ws_frango, ws_prod, excel)
        atualizar_apinco(ws_frango, ws_prod, excel)

        excel.Calculation = XL_CALC_AUTO
        wb.Save()
        wb.Close(SaveChanges=False)
        print(f"\nSalvo: {path.name}")

    except Exception:
        import traceback
        traceback.print_exc()
        sys.exit(1)
    finally:
        try:
            excel.ScreenUpdating = True
            excel.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
