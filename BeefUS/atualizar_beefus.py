"""
atualizar_beefus.py (versao openpyxl / nuvem)

Consolidador do BeefUS. Estende formulas e formatos sem Excel, gravando o
BeefUS.xlsm via Google Drive (drive_io). As formulas novas so ganham cache de
valor depois do Excel/Bloomberg recalcular a planilha.
"""

from __future__ import annotations

import datetime as dt
import re
import sys
from copy import copy
from pathlib import Path

import openpyxl

XLSX_NAME = "BeefUS.xlsm"
SHEET_MAIN = "BeefUS"
SHEET_ABATES = "Abates"

MAIN_FIRST_DATA_ROW = 5
ABATES_FIRST_DATA_ROW = 3
ROW_OFFSET_ABATES_TO_MAIN = MAIN_FIRST_DATA_ROW - ABATES_FIRST_DATA_ROW

ABATES_MAIN_COL_START = 4
COL_DATE_MAIN = 2
COL_M = 13
COL_N = 14
COL_O = 15
COL_P = 16
COL_Q = 17

_REF = re.compile(r"(?<![A-Za-z0-9_])(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z_(])")


def shift_formula(formula, delta: int):
    if not isinstance(formula, str) or not formula.startswith("="):
        return formula

    def repl(m):
        col_abs, col, row_abs, row = m.groups()
        if row_abs == "$":
            return m.group(0)
        return f"{col_abs}{col}{row_abs}{int(row) + delta}"

    return _REF.sub(repl, formula)


def copy_cell(src, dst, *, shift_delta: int = 0) -> None:
    src_value = src.value
    dst.value = shift_formula(src_value, shift_delta) if isinstance(src_value, str) and src_value.startswith("=") else src_value
    dst.number_format = src.number_format
    dst.font = copy(src.font)
    dst.fill = copy(src.fill)
    dst.border = copy(src.border)
    dst.alignment = copy(src.alignment)
    dst.protection = copy(src.protection)


def extend_section(ws, src_row: int, col_ranges: list[tuple[int, int]], dst_end: int) -> int:
    if dst_end <= src_row:
        return 0
    for c0, c1 in col_ranges:
        for col in range(c0, c1 + 1):
            src = ws.cell(row=src_row, column=col)
            for dst_row in range(src_row + 1, dst_end + 1):
                copy_cell(src, ws.cell(row=dst_row, column=col), shift_delta=dst_row - src_row)
    return dst_end - src_row


def last_value_row(ws, col: int, start: int) -> int:
    for row in range(ws.max_row, start - 1, -1):
        value = ws.cell(row=row, column=col).value
        if value not in (None, ""):
            return row
    return start - 1


def last_formula_row(ws, col: int, start: int) -> int:
    for row in range(ws.max_row, start - 1, -1):
        value = ws.cell(row=row, column=col).value
        if isinstance(value, str) and value.startswith("="):
            return row
    return start - 1


def row_for_year_month(ws, year: int, month: int) -> int | None:
    for row in range(MAIN_FIRST_DATA_ROW, ws.max_row + 1):
        value = ws.cell(row=row, column=COL_DATE_MAIN).value
        if isinstance(value, dt.datetime):
            value = value.date()
        if isinstance(value, dt.date) and value.year == year and value.month == month:
            return row
    return None


def sincronizar_abates(ws_main, ws_abates) -> dict:
    last_abates = last_value_row(ws_abates, col=3, start=ABATES_FIRST_DATA_ROW)
    if last_abates < ABATES_FIRST_DATA_ROW:
        return {"mudou": False, "msg": "Abates: nenhuma linha com dado em col C."}

    target_main = last_abates + ROW_OFFSET_ABATES_TO_MAIN
    template_main = last_formula_row(ws_main, col=ABATES_MAIN_COL_START, start=MAIN_FIRST_DATA_ROW)
    if template_main < MAIN_FIRST_DATA_ROW:
        return {"mudou": False, "msg": "BeefUS: nenhuma linha-template com formula em D."}
    if target_main <= template_main:
        return {"mudou": False, "msg": f"Abates: ja alinhado (ultima linha BeefUS={template_main})."}

    n = extend_section(ws_main, template_main, [(1, COL_P)], target_main)
    return {"mudou": True, "msg": f"Abates: +{n} linha(s) em BeefUS (ate {target_main})."}


def ajustar_precos(ws_main, today: dt.date) -> dict:
    row_cur = row_for_year_month(ws_main, today.year, today.month)
    if row_cur is None:
        return {"mudou": False, "msg": "Precos: nao encontrei a linha do mes atual na col B."}

    prev = today.replace(day=1) - dt.timedelta(days=1)
    row_prev = row_for_year_month(ws_main, prev.year, prev.month) or row_cur - 1
    if row_prev < MAIN_FIRST_DATA_ROW:
        return {"mudou": False, "msg": "Precos: sem linha anterior para usar como template."}

    mudou = False
    acoes: list[str] = []

    for col, label in [(COL_M, "M"), (COL_Q, "Q")]:
        last_row = last_formula_row(ws_main, col=col, start=MAIN_FIRST_DATA_ROW)
        if last_row >= MAIN_FIRST_DATA_ROW and last_row < row_cur:
            extend_section(ws_main, last_row, [(col, col)], row_cur)
            mudou = True
            acoes.append(f"{label}: formula copiada linhas {last_row + 1}->{row_cur} ({row_cur - last_row} mes(es))")

    n_val = ws_main.cell(row=row_cur, column=COL_N).value
    if n_val not in (None, ""):
        for col, label in [(COL_O, "O"), (COL_P, "P")]:
            cell = ws_main.cell(row=row_cur, column=col)
            has_formula = isinstance(cell.value, str) and cell.value.startswith("=")
            if not has_formula and cell.value in (None, ""):
                copy_cell(ws_main.cell(row=row_prev, column=col), cell, shift_delta=row_cur - row_prev)
                mudou = True
                acoes.append(f"{label}: formula copiada {row_prev}->{row_cur} (N preenchida)")

    if not mudou:
        return {"mudou": False, "msg": "Precos: nada para ajustar."}
    return {"mudou": True, "msg": "Precos: " + "; ".join(acoes)}


def main() -> None:
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    import drive_io

    fid = drive_io.find_file_id(XLSX_NAME)
    path = drive_io.download_to_temp(fid)
    today = dt.date.today()
    print(f"[{dt.datetime.now():%Y-%m-%d %H:%M}] BeefUS consolidacao - {today:%d/%m/%Y}")

    wb = openpyxl.load_workbook(path, keep_vba=True, data_only=False)
    try:
        ws_main = wb[SHEET_MAIN]
        ws_abates = wb[SHEET_ABATES]

        r1 = sincronizar_abates(ws_main, ws_abates)
        print(" - " + r1["msg"])
        r2 = ajustar_precos(ws_main, today)
        print(" - " + r2["msg"])

        if r1.get("mudou") or r2.get("mudou"):
            wb.save(path)
            drive_io.update_file(fid, path)
            print("Salvo. Formulas novas serao calculadas no PC Bloomberg.")
        else:
            print("Nada alterado - nao salvou.")
    finally:
        wb.close()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        raise SystemExit(130)
