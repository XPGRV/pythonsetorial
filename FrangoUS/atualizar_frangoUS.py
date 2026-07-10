"""
atualizar_frangoUS.py

FrangoUS — atualiza formulas em duas abas:

Aba BBG_Dados:
  Colunas I, J, L, O, Q, R, V, X, AA — ultima linha com dado em col E.

Aba FrangoUS:
  (i)   Mes corrente:  D->F, I->P, W, AP, AQ  — ate o mes corrente
  (ii)  Mes anterior:  AE->AH, AK->AN, AR      — sempre o mes anterior ao corrente
  (iii) Prod col C:    Z                        — ultimo dado em Production col C
  (iv)  Prod col D:    AA->AD                   — ultimo dado em Production col D
  (v)   Prod col G:    AI, AJ                   — ultimo dado em Production col G
  (vi)  Prod col H:    AO                       — ultimo dado em Production col H

Mapeamento de linhas:
  FrangoUS:   linha 5 = jan/2000  ->  row = 5 + (ano-2000)*12 + (mes-1)
  Production: linha 3 = jan/2000  ->  FrangoUS_row = Production_row + 2

Requer: pip install pywin32
"""

import sys
import datetime
from pathlib import Path

DRIVE_ROOT = Path(
    r"G:\.shortcut-targets-by-id\11q1ngdqbySnDMCQSP4fiwPztT6ju2F4Z"
)
XLSX_NAME = "FrangoUS.xlsm"

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


def frangoUS_row(year: int, month: int) -> int:
    """Linha no FrangoUS para ano/mes. Linha 5 = jan/2000."""
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


def last_data_row(ws, col: int, start: int = 3) -> int:
    """Ultima linha em 'col' com valor numerico > 0, varrendo de baixo para cima."""
    bottom = ws.Cells(ws.Rows.Count, col).End(XL_UP).Row
    for r in range(bottom, start - 1, -1):
        v = ws.Cells(r, col).Value
        if isinstance(v, (int, float)) and v > 0:
            return r
    return 0


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
    """Cols I, J, L, O, Q, R, V, X, AA — ultima linha com dado em col E."""
    last_data = ws_bbg.Cells(ws_bbg.Rows.Count, 5).End(XL_UP).Row
    last_tmpl = last_formula_row(ws_bbg, col=9, start=2)  # col I

    if last_tmpl < 2:
        print("  BBG_Dados: nenhuma linha template em col I — pulando.")
        return

    if last_data > last_tmpl:
        n = extend_section(
            ws_bbg, last_tmpl,
            [(9, 10), (12, 12), (15, 15), (17, 18), (22, 22), (24, 24), (27, 27)],
            # I,J | L | O | Q,R | V | X | AA
            last_data, excel
        )
        print(f"  BBG_Dados: +{n} linha(s) -> linha {last_data}")
    else:
        print(f"  BBG_Dados: ja atualizado (linha {last_tmpl})")


# ── FrangoUS ──────────────────────────────────────────────────────────────────

def atualizar_mes_corrente(ws_us, today: datetime.date, excel) -> None:
    """(i) Cols D->F, I->P, W, AP, AQ — ate o mes corrente."""
    target = frangoUS_row(today.year, today.month)
    last   = last_formula_row(ws_us, col=4)  # col D

    if last < 5:
        print("  (i)  Mes corrente: nenhuma linha template em col D — pulando.")
        return

    if target > last:
        n = extend_section(
            ws_us, last,
            [(4, 6), (9, 16), (23, 23), (42, 43)],  # D->F | I->P | W | AP,AQ
            target, excel
        )
        print(f"  (i)  Mes corrente: +{n} mes(es) -> {format_ym(today.year, today.month)}")
    else:
        print(f"  (i)  Mes corrente: ja atualizado (linha {last})")


def atualizar_mes_anterior(ws_us, today: datetime.date, excel) -> None:
    """(ii) Cols AE->AH, AK->AN, AR — sempre o mes anterior ao corrente."""
    yr, mo = prev_month(today.year, today.month)
    target = frangoUS_row(yr, mo)
    last   = last_formula_row(ws_us, col=31)  # col AE

    if last < 5:
        print("  (ii) Mes anterior: nenhuma linha template em col AE — pulando.")
        return

    if target > last:
        n = extend_section(
            ws_us, last,
            [(31, 34), (37, 40), (44, 44)],  # AE->AH | AK->AN | AR
            target, excel
        )
        print(f"  (ii) Mes anterior: +{n} mes(es) -> {format_ym(yr, mo)}")
    else:
        print(f"  (ii) Mes anterior: ja atualizado (linha {last})")


def atualizar_prod_col(ws_us, ws_prod, prod_col: int, frango_col_ranges: list,
                       label: str, excel) -> None:
    """Generico: estende colunas de ws_us ate o ultimo dado em ws_prod[prod_col]."""
    last_prod = last_data_row(ws_prod, col=prod_col)

    if last_prod == 0:
        print(f"  {label}: sem dados em Production col {col_letter(prod_col)} — pulando.")
        return

    target    = last_prod + 2  # Production linha 3 = FrangoUS linha 5
    ref_col   = frango_col_ranges[0][0]
    last_tmpl = last_formula_row(ws_us, col=ref_col)

    if last_tmpl < 5:
        print(f"  {label}: nenhuma linha template em col {col_letter(ref_col)} — pulando.")
        return

    if target > last_tmpl:
        yr, mo = production_row_to_ym(last_prod)
        n = extend_section(ws_us, last_tmpl, frango_col_ranges, target, excel)
        print(f"  {label}: +{n} mes(es) -> {format_ym(yr, mo)}")
    else:
        print(f"  {label}: ja atualizado (linha {last_tmpl})")


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
    print(f"FrangoUS — {today:%d/%m/%Y}\n")

    excel = win32com.client.DispatchEx("Excel.Application")
    try: excel.Visible        = False
    except Exception: pass
    try: excel.DisplayAlerts  = False
    except Exception: pass
    try: excel.ScreenUpdating = False
    except Exception: pass

    try:
        wb      = excel.Workbooks.Open(str(path))
        excel.Calculation = XL_CALC_MANUAL
        ws_bbg  = wb.Sheets("BBG_Dados")
        ws_us   = wb.Sheets("FrangoUS")
        ws_prod = wb.Sheets("Production")

        atualizar_bbg_dados(ws_bbg, excel)
        atualizar_mes_corrente(ws_us, today, excel)
        atualizar_mes_anterior(ws_us, today, excel)

        # Prod col C -> Z (26)
        atualizar_prod_col(ws_us, ws_prod, prod_col=3,
                           frango_col_ranges=[(26, 26)],
                           label="(iii) Prod C -> Z", excel=excel)

        # Prod col D -> AA->AD (27->30)
        atualizar_prod_col(ws_us, ws_prod, prod_col=4,
                           frango_col_ranges=[(27, 30)],
                           label="(iv)  Prod D -> AA->AD", excel=excel)

        # Prod col G -> AI, AJ (35->36)
        atualizar_prod_col(ws_us, ws_prod, prod_col=7,
                           frango_col_ranges=[(35, 36)],
                           label="(v)   Prod G -> AI,AJ", excel=excel)

        # Prod col H -> AO (41)
        atualizar_prod_col(ws_us, ws_prod, prod_col=8,
                           frango_col_ranges=[(41, 41)],
                           label="(vi)  Prod H -> AO", excel=excel)

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
