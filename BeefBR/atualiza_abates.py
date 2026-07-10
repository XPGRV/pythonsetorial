"""Atualizacao mensal SIGSIF - aba Abates de BeefBR.xlsm, via Drive."""
from __future__ import annotations
import datetime as dt
import io
import sys
from copy import copy
from pathlib import Path
import openpyxl
import pandas as pd
import requests

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import drive_io

XLSX_NAME = "BeefBR.xlsm"
SHEET_NAME = "Abates"
ANO_INICIO = 2021
SIF_URL = "https://dados.agricultura.gov.br/dataset/062166e3-b515-4274-8e7d-68aadd64b820/resource/341dc717-4716-42ab-b189-c8d7a9d2a1ba/download/sigsifrelatorioabates.csv"
CATS = {"Bovino", "Touro/Touruno", "Vaca", "Novilho Precoce", "Novilhao", "Novilho Intermediario", "Novilha Precoce", "Novilha Intermediaria", "Novilhona", "Vitelo"}


def row_of(year, month):
    return 3 + (year - 1997) * 12 + month - 1


def shift_formula(formula, delta):
    import re
    if not isinstance(formula, str) or not formula.startswith("="):
        return formula
    return re.sub(r"(?<![A-Za-z0-9_])(\$?[A-Za-z]{1,3})(\$?)(\d+)", lambda m: m.group(0) if m.group(2) == "$" else f"{m.group(1)}{m.group(2)}{int(m.group(3)) + delta}", formula)


def fetch_data():
    # O servidor do MAPA (CKAN) e instavel; o retry automatico (net.py, via
    # drive_io) repete a chamada em caso de timeout/erro transitorio.
    # timeout=(connect, read): conecta rapido ou desiste em 15s (e ai o retry
    # tenta de novo, em vez de travar 3min); leitura do CSV pode levar ate 120s.
    response = requests.get(SIF_URL, timeout=(15, 120))
    response.raise_for_status()
    frame = None
    for encoding in ("utf-8", "latin-1", "cp1252"):
        try:
            frame = pd.read_csv(io.BytesIO(response.content), sep=";", encoding=encoding, dtype=str)
            break
        except UnicodeDecodeError:
            pass
    if frame is None:
        raise RuntimeError("Nao foi possivel ler o CSV SIGSIF.")
    frame.columns = frame.columns.str.strip()
    frame["CATEGORIA"] = frame["CATEGORIA"].str.strip()
    frame["date"] = pd.to_datetime(frame["MES_ANO"].str.strip(), format="%m/%Y", errors="coerce")
    frame = frame[(frame.date.notna()) & (frame.date.dt.year >= ANO_INICIO) & frame.CATEGORIA.isin(CATS)].copy()
    frame["QTD_MACHO"] = pd.to_numeric(frame["QTD_MACHO"], errors="coerce").fillna(0)
    frame["QTD_FEMEA"] = pd.to_numeric(frame["QTD_FEMEA"], errors="coerce").fillna(0)
    grouped = frame.groupby([frame.date.dt.year, frame.date.dt.month])[["QTD_MACHO", "QTD_FEMEA"]].sum()
    return {(int(y), int(m)): (int(r.QTD_MACHO), int(r.QTD_FEMEA)) for (y, m), r in grouped.iterrows()}


def copy_template(ws, source_row, target_row):
    for col in range(22, 27):
        source, target = ws.cell(source_row, col), ws.cell(target_row, col)
        target._style = copy(source._style)
        target.number_format = source.number_format
        target.font = copy(source.font)
        target.fill = copy(source.fill)
        target.border = copy(source.border)
        target.alignment = copy(source.alignment)
        target.value = shift_formula(source.value, target_row - source_row)


def main():
    data = fetch_data()
    fid = drive_io.find_file_id(XLSX_NAME)
    path = drive_io.download_to_temp(fid)
    wb = openpyxl.load_workbook(path, keep_vba=True, data_only=False)
    try:
        ws = wb[SHEET_NAME]
        template = next((r for r in range(ws.max_row, 2, -1) if any(ws.cell(r, c).value not in (None, "") for c in range(3, 8))), None)
        changes = 0
        for (year, month), (male, female) in sorted(data.items()):
            row = row_of(year, month)
            if row > ws.max_row:
                continue
            new_row = ws.cell(row, 22).value in (None, "") and ws.cell(row, 23).value in (None, "")
            if new_row and template:
                copy_template(ws, template, row)
            for col, value in ((22, male), (23, female)):
                if ws.cell(row, col).value != value:
                    ws.cell(row, col).value = value
                    changes += 1
        if changes:
            wb.save(path)
            drive_io.update_file(fid, path)
        print(f"SIGSIF BeefBR: {changes} celula(s) atualizada(s).")
    finally:
        wb.close()


if __name__ == "__main__":
    main()
