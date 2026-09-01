"""Atualiza SECEX Brasil e o cambio diario do BeefBR.xlsm via Drive."""
from __future__ import annotations

import datetime as dt
import sys
import time
from copy import copy
from pathlib import Path

import openpyxl
import requests

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import drive_io

XLSX_NAME = "BeefBR.xlsm"
SHEET_NAME = "SECEX"
COMEX_URL = "https://api-comexstat.mdic.gov.br/general"
BCB_URL = "https://api.bcb.gov.br/dados/serie/bcdata.sgs.1/dados"
CODES = ("0203", "0201", "0202", "0207")


def month_key(value):
    if isinstance(value, dt.datetime):
        value = value.date()
    if isinstance(value, dt.date):
        return value.year, value.month
    return None


def as_number(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def comex(code, start_year):
    payload = {
        "flow": "export",
        "monthDetail": True,
        "period": {"from": f"{start_year}-01", "to": f"{dt.date.today().year}-12"},
        "filters": [{"filter": "heading", "values": [code]}],
        "details": ["heading"],
        "metrics": ["metricFOB", "metricKG"],
    }
    for attempt in range(4):
        response = requests.post(
            COMEX_URL,
            json=payload,
            headers={"Content-Type": "application/json"},
            timeout=90,
        )
        if response.status_code == 429 and attempt < 3:
            time.sleep(25)
            continue
        response.raise_for_status()
        return response.json().get("data", {}).get("list", [])
    return []


def bcb_fx(start_date, end_date):
    response = requests.get(
        BCB_URL,
        params={
            "formato": "json",
            "dataInicial": start_date.strftime("%d/%m/%Y"),
            "dataFinal": end_date.strftime("%d/%m/%Y"),
        },
        timeout=90,
    )
    # O SGS devolve 404 -- nao 200 com lista vazia -- quando o intervalo pedido
    # nao tem cotacao publicada. Acontece todo dia util antes de ~13h (a PTAX do
    # dia ainda nao saiu) e em fim de semana e feriado, porque nesses casos o
    # intervalo vira "hoje ate hoje". Nao e erro: e so nao ter FX novo ainda.
    if response.status_code == 404:
        print(
            f"[BCB] sem cotacao de {start_date:%d/%m/%Y} a {end_date:%d/%m/%Y}; "
            "seguindo sem FX novo nesta rodada.",
            flush=True,
        )
        return {}
    response.raise_for_status()
    rates = {}
    for item in response.json():
        try:
            date = dt.datetime.strptime(item["data"], "%d/%m/%Y").date()
            # BCB (formato=json) usa ponto decimal, ex "5.1329". Nao mexer no ponto!
            rates[date] = float(item["valor"])
        except (KeyError, TypeError, ValueError):
            continue
    return rates


def fx_formula(row):
    """Formula AVERAGEIFS: media mensal do FX diario (BBG_Dados col G) do mes em
    B{row}. Igual as linhas antigas da SECEX; o Excel/Bloomberg calcula o valor."""
    return (
        f'=AVERAGEIFS(BBG_Dados!G:G,BBG_Dados!D:D,'
        f'">="&DATE(YEAR(SECEX!B{row}),MONTH(SECEX!B{row}),1),'
        f'BBG_Dados!D:D,"<"&DATE(YEAR(SECEX!B{row}),MONTH(SECEX!B{row})+1,1))'
    )


def copy_style(ws, source_row, target_row, columns):
    for column in columns:
        source = ws.cell(source_row, column)
        target = ws.cell(target_row, column)
        target._style = copy(source._style)
        target.number_format = source.number_format
        target.font = copy(source.font)
        target.fill = copy(source.fill)
        target.border = copy(source.border)
        target.alignment = copy(source.alignment)


def update_daily_fx(workbook, rates, today):
    ws = workbook["BBG_Dados"]
    changes = 0
    for row in range(4, ws.max_row + 1):
        value = ws.cell(row, 4).value
        if isinstance(value, dt.datetime):
            value = value.date()
        if not isinstance(value, dt.date) or value > today:
            continue
        rate = rates.get(value)
        if rate is None:
            continue
        cell = ws.cell(row, 7)
        # Bloomberg continua sendo a fonte quando ja houver observacao;
        # o BCB preenche apenas os dias que ainda nao chegaram pelo terminal.
        if as_number(cell.value) is not None:
            continue
        cell.value = rate
        if row > 4:
            cell._style = copy(ws.cell(row - 1, 7)._style)
        changes += 1
    return changes


def main():
    today = dt.date.today()
    probe = comex("0203", max(2000, today.year - 2))
    if not probe:
        raise RuntimeError("Comex Stat nao retornou meses para determinar a ultima publicacao.")
    last_api = max((int(item["year"]), int(item["monthNumber"])) for item in probe)

    file_id = drive_io.find_file_id(XLSX_NAME)
    path = drive_io.download_to_temp(file_id)
    workbook = openpyxl.load_workbook(path, keep_vba=True, data_only=False)
    try:
        ws = workbook[SHEET_NAME]
        rows = {
            month_key(ws.cell(row, 2).value): row
            for row in range(3, ws.max_row + 1)
            if month_key(ws.cell(row, 2).value)
        }
        eligible = sorted(period for period in rows if period <= last_api)
        # O Comex pode revisar meses ja publicados. Reconciliar todo o historico
        # publicado corrige volume e precos sem avancar para meses ainda indisponiveis.
        targets = eligible

        fx_missing = [
            period for period, row in rows.items()
            if period <= (today.year, today.month) and ws.cell(row, 15).value in (None, "")
        ]
        fx_period = min(fx_missing) if fx_missing else (today.year, today.month)
        fx_start = dt.date(fx_period[0], fx_period[1], 1)
        rates = bcb_fx(fx_start, today)
        fx_changes = update_daily_fx(workbook, rates, today)

        raw = {code: {} for code in CODES}
        if targets:
            start_year = min(year for year, _ in targets)
            for code in CODES:
                for item in comex(code, start_year):
                    period = int(item["year"]), int(item["monthNumber"])
                    raw[code][period] = (
                        float(item.get("metricKG") or 0),
                        float(item.get("metricFOB") or 0),
                    )
                time.sleep(10)

        changes = 0
        for period in targets:
            row = rows[period]
            pork = raw["0203"].get(period, (0, 0))
            fresh = raw["0201"].get(period, (0, 0))
            frozen = raw["0202"].get(period, (0, 0))
            poultry = raw["0207"].get(period, (0, 0))
            beef = fresh[0] + frozen[0], fresh[1] + frozen[1]
            if not any((pork[0], beef[0], poultry[0])):
                continue

            reference = max(
                (candidate for candidate in range(3, row) if as_number(ws.cell(candidate, 4).value) is not None),
                default=None,
            )
            if reference:
                copy_style(ws, reference, row, range(4, 16))

            values = {
                4: pork[0] / 1e6 if pork[0] else None,
                5: beef[0] / 1e6 if beef[0] else None,
                6: poultry[0] / 1e6 if poultry[0] else None,
                8: pork[1] / pork[0] if pork[0] else None,
                10: beef[1] / beef[0] if beef[0] else None,
                12: poultry[1] / poultry[0] if poultry[0] else None,
            }
            # O (FX) e a media mensal do FX diario -> FORMULA AVERAGEIFS (igual as
            # linhas antigas). I/K/M (BRL/Kg) = USD/kg * FX -> tambem formulas.
            # (Nao hardcodar: o Excel/Bloomberg calcula ao recalcular.)
            values[15] = fx_formula(row)
            if values[8] is not None:
                values[9] = f"=H{row}*$O{row}"
            if values[10] is not None:
                values[11] = f"=J{row}*$O{row}"
            if values[12] is not None:
                values[13] = f"=L{row}*$O{row}"
            for column, value in values.items():
                if value is None:
                    continue
                ws.cell(row, column).value = value
                changes += 1

        # O cambio mensal deve existir ate hoje, mesmo quando o SECEX ainda nao publicou o mes.
        for period, row in rows.items():
            if period > (today.year, today.month):
                continue
            cell = ws.cell(row, 15)
            # Vazio OU valor hardcoded (bug antigo do x10000) -> poe a formula.
            if not (isinstance(cell.value, str) and cell.value.startswith("=")):
                cell.value = fx_formula(row)
                if row > 3:
                    cell._style = copy(ws.cell(row - 1, 15)._style)
                changes += 1

        if changes or fx_changes:
            workbook.calculation.fullCalcOnLoad = True
            workbook.calculation.forceFullCalc = True
            workbook.save(path)
            drive_io.update_file(file_id, path)
        print(
            f"SECEX BeefBR: ultimo mes Comex {last_api[1]:02d}/{last_api[0]}; "
            f"{changes} celula(s) SECEX/cambio e {fx_changes} dia(s) de FX atualizados."
        )
    finally:
        workbook.close()


if __name__ == "__main__":
    main()