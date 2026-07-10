"""
atualizar_beefbr.py

Atualiza as formulas da aba BeefBR em tres secoes:
  (i)   Spreads    D->S (exceto E)  — ate o mes corrente
  (ii)  IBGE       V->AC            — ate o ultimo mes com dado na aba Abates (IBGE, cols C->M)
  (iii) SIF        AF->AL           — ate o ultimo mes FECHADO na aba Abates (SIF, cols V->Z)

Requer: pip install pywin32
"""

import sys
import datetime
from pathlib import Path

DRIVE_ROOT = Path(
    r"G:\.shortcut-targets-by-id\11q1ngdqbySnDMCQSP4fiwPztT6ju2F4Z"
)
XLSX_NAME = "BeefBR.xlsm"

# Constantes Excel via win32com
XL_UP             = -4162
XL_PASTE_FORMULAS = -4123
XL_PASTE_ALL      = -4104   # formulas + formatacao completa
XL_CALC_MANUAL    = -4135
XL_CALC_AUTO      = -4105


# ── Helpers ──────────────────────────────────────────────────────────────────

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
    """Numero de coluna (1-based) para letra Excel (1=A, 27=AA...)."""
    result = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def beefbr_row(year: int, month: int) -> int:
    """Linha no BeefBR para ano/mes. Linha 5 = jan/2000."""
    return 5 + (year - 2000) * 12 + (month - 1)


def last_formula_row(ws, col: int, start: int = 5, array_only: bool = False) -> int:
    """Ultima linha em 'col' com formula, buscando de baixo para cima (End xlUp)."""
    last = ws.Cells(ws.Rows.Count, col).End(XL_UP).Row
    if last < start:
        return start - 1
    for r in range(last, start - 1, -1):
        cell = ws.Cells(r, col)
        has_it = cell.HasArray if array_only else cell.HasFormula
        if has_it:
            return r
    return start - 1


def extend_section(ws, src_row: int, col_ranges: list, dst_end: int, excel) -> int:
    """
    Copia formulas de src_row para as linhas src_row+1 ate dst_end,
    nas faixas de colunas em col_ranges [(c_ini, c_fim), ...].
    Retorna o numero de linhas novas.
    """
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


def prev_month(year: int, month: int):
    return (year - 1, 12) if month == 1 else (year, month - 1)


def format_date(raw) -> str:
    try:
        return f"{raw.month:02d}/{raw.year}"
    except Exception:
        return str(raw)


# ── Logica por secao ──────────────────────────────────────────────────────────

def atualizar_spreads(ws_beef, today: datetime.date, excel) -> None:
    """
    (i) Spreads — dois grupos:
      Parte 1: D, G->M (pula E, F e J) + AN  — ate o mes corrente.
      Parte 2: F e N->S (Preco/Spread ME)    — ate o ultimo mes com dado em col E (Preco ME).
    Template: ultima linha onde col G tem formula.
    """
    target   = beefbr_row(today.year, today.month)
    last_g   = last_formula_row(ws_beef, col=7,  array_only=False)  # col G
    last_an  = last_formula_row(ws_beef, col=40, array_only=False)  # col AN

    if last_g < 5:
        print("  (i)  Spreads: nenhuma linha template em col G — pulando.")
        return
    if last_an < 5:
        print("  (i)  Spreads/AN: nenhuma linha template em col AN — pulando AN.")
        last_an = last_g  # nao bloqueia o resto

    last = min(last_g, last_an)

    # ── Parte 1: D, G->M, AN — ate o mes corrente ────────────────────────────
    if target > last:
        n = extend_section(
            ws_beef, last,
            [(4, 4), (7, 9), (11, 13), (40, 40)],  # col D | G->I | K->M | AN
            target, excel
        )
        print(f"  (i)  Spreads: +{n} mes(es) -> {format_date(ws_beef.Cells(target, 2).Value)}")
    else:
        print(f"  (i)  Spreads: ja atualizado (linha {last})")

    # ── Parte 2: F e N->S (Preco/Spread ME) — ate ultimo mes com dado em col E ─
    last_f = last_formula_row(ws_beef, col=6,  array_only=False)  # col F
    last_n = last_formula_row(ws_beef, col=14, array_only=False)  # col N
    if last_f < 5 and last_n < 5:
        print("  (i)  Preco/Spread ME (F, N->S): nenhuma linha template — pulando.")
        return
    last_fn = min(r for r in (last_f, last_n) if r >= 5)

    # Busca de baixo para cima a ultima linha de col E com valor numerico
    target_f = last_fn  # fallback: nao extende se col E nao tiver dado novo
    bottom_f = ws_beef.Cells(ws_beef.Rows.Count, 5).End(XL_UP).Row
    for r in range(bottom_f, 4, -1):
        v = ws_beef.Cells(r, 5).Value
        try:
            if v is not None and float(v) > 0:
                target_f = r
                break
        except (TypeError, ValueError):
            pass

    if target_f > last_fn:
        n2 = extend_section(
            ws_beef, last_fn,
            [(6, 6), (14, 19)],  # F | N->S
            target_f, excel
        )
        print(f"  (i)  Preco/Spread ME (F, N->S): +{n2} mes(es) -> {format_date(ws_beef.Cells(target_f, 2).Value)}")
    else:
        print(f"  (i)  Preco/Spread ME (F, N->S): ja atualizado (linha {last_fn})")


def atualizar_ibge(ws_beef, ws_abates, excel) -> None:
    """
    (ii) IBGE Abates: colunas V->AC.
    Alvo = ultimo mes com dado bruto na aba Abates, col C (Bois).
    Offset: BeefBR_row = Abates_row - 34.
    Template: ultima linha onde col V (22) tem formula.
    """
    # Ultimo dado IBGE em Abates (col C = 3, dado bruto > 0)
    last_abates = 38
    consecutive_empty = 0
    for r in range(39, 500):
        v = ws_abates.Cells(r, 3).Value
        if isinstance(v, (int, float)) and v > 0:
            last_abates = r
            consecutive_empty = 0
        else:
            consecutive_empty += 1
            if consecutive_empty > 12:
                break

    target = last_abates - 34
    last   = last_formula_row(ws_beef, col=22)  # col V

    if last < 5:
        print("  (ii) IBGE Abates: nenhuma linha template em col V — pulando.")
        return

    if target > last:
        n = extend_section(
            ws_beef, last,
            [(22, 29)],          # cols V->AC
            target, excel
        )
        print(f"  (ii) IBGE Abates: +{n} mes(es) -> {format_date(ws_beef.Cells(target, 2).Value)}")
    else:
        print(f"  (ii) IBGE Abates: ja atualizado (linha {last})")


def atualizar_sif(ws_beef, ws_abates, excel) -> None:
    """
    (iii) SIF Abates: colunas AF->AL.
    Alvo = 2 meses antes do ultimo mes com dado SIF na aba Abates (col V).
    Ex: se Abates col V vai ate mai/2026, BeefBR vai ate mar/2026.
    Template: ultima linha onde col AF (32) tem formula.
    """
    bottom = ws_abates.Cells(ws_abates.Rows.Count, 22).End(XL_UP).Row
    last_abates = 0
    for r in range(bottom, 37, -1):
        v = ws_abates.Cells(r, 22).Value
        if isinstance(v, (int, float)) and v > 0:
            last_abates = r
            break

    if last_abates == 0:
        print("  (iii) SIF Abates: sem dados em col V da aba Abates — pulando.")
        return

    raw_date = ws_abates.Cells(last_abates, 2).Value
    if raw_date is None:
        print("  (iii) SIF Abates: data ilegivel em Abates — pulando.")
        return

    try:
        yr, mo = raw_date.year, raw_date.month
    except Exception:
        print("  (iii) SIF Abates: data ilegivel em Abates — pulando.")
        return

    yr, mo = prev_month(yr, mo)
    yr, mo = prev_month(yr, mo)
    target = beefbr_row(yr, mo)
    last   = last_formula_row(ws_beef, col=32)  # col AF

    if last < 5:
        print("  (iii) SIF Abates: nenhuma linha template em col AF — pulando.")
        return

    if target > last:
        n = extend_section(
            ws_beef, last,
            [(32, 38)],
            target, excel
        )
        print(f"  (iii) SIF Abates: +{n} mes(es) -> {format_date(ws_beef.Cells(target, 2).Value)}")
    else:
        print(f"  (iii) SIF Abates: ja atualizado (linha {last})")


def atualizar_bbg_dados(ws_bbg, excel) -> None:
    """
    BBG_Dados: colunas I e J alinhadas com a ultima linha com dado em cols E->H.
    Usa col G (USDBRL) como referencia — FX disponivel em todos os dias uteis.
    """
    # Ultima linha com dado em E->H (referencia: col G = USDBRL)
    last_data = ws_bbg.Cells(ws_bbg.Rows.Count, 7).End(XL_UP).Row

    # Template: ultima linha com formula em col I (start=2, header na linha 1)
    last_ij = last_formula_row(ws_bbg, col=9, start=2)

    if last_ij < 2:
        print("  (iv) BBG_Dados I->J: nenhuma linha template em col I — pulando.")
        return

    if last_data > last_ij:
        n = extend_section(ws_bbg, last_ij, [(9, 10)], last_data, excel)
        print(f"  (iv) BBG_Dados I->J: +{n} linha(s)")
    else:
        print(f"  (iv) BBG_Dados I->J: ja atualizado (linha {last_ij})")


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    try:
        import win32com.client
    except ImportError:
        print("ERRO: pywin32 nao instalado. Execute: pip install pywin32")
        sys.exit(1)

    try:
        path = localizar_planilha()
    except (FileNotFoundError, RuntimeError) as exc:
        print(f"ERRO: {exc}")
        sys.exit(1)

    today = datetime.date.today()
    print(f"BeefBR — {today:%d/%m/%Y}\n")

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
        ws_beef   = wb.Sheets("BeefBR")
        ws_abates = wb.Sheets("Abates")
        ws_bbg    = wb.Sheets("BBG_Dados")

        atualizar_bbg_dados(ws_bbg, excel)
        atualizar_spreads(ws_beef, today, excel)
        atualizar_ibge(ws_beef, ws_abates, excel)
        atualizar_sif(ws_beef, ws_abates, excel)

        excel.Calculation = XL_CALC_AUTO
        wb.Save()
        wb.Close(SaveChanges=False)
        print(f"\nSalvo: {path.name}")

    except Exception as e:
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
