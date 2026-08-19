#!/usr/bin/env python3
"""
Atualiza os dados SECEX de transformadores na planilha WEG - Setorial.xlsm.

Fonte: Comex Stat API
Fluxo: exportacao
Detalhe: mes + UF do produto
Valor: US$ FOB, gravado em 1000 US$
Preco: US$/unid ou US$/kg, sem abertura por estado

Planilha:
  G:\\.shortcut-targets-by-id\\11q1ngdqbySnDMCQSP4fiwPztT6ju2F4Z\\Arquivos\\Setorial - Database\\WEG - Setorial.xlsm
  Aba: Transformadores

Comportamento:
  - Escreve apenas celulas vazias.
  - Nao revisa nem sobrescreve valores ja preenchidos.
  - Consulta a Comex Stat apenas para produtos/campos com celulas vazias.
  - Para cada produto, consulta a partir do primeiro ano vazio daquele produto.

Uso:
  python update_secex_transformadores.py
  python update_secex_transformadores.py --dry-run
"""

from __future__ import annotations

import argparse
import sys
import time
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

import requests
import openpyxl
from openpyxl.utils import column_index_from_string

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import drive_io


sys.stdout.reconfigure(encoding="utf-8")


WORKBOOK_NAME = "WEG - Setorial.xlsm"   # localizada no Drive via drive_io
SHEET_NAME = "Transformadores"
FIRST_DATA_ROW = 5

API_URL = "https://api-comexstat.mdic.gov.br/general"
API_HEADERS = {"Content-Type": "application/json"}

WAIT_NORMAL = 8
WAIT_429 = 25
MAX_RETRY = 4

VALUE_KEYS = ("total", "sc")
MISSING_KEYS = VALUE_KEYS + ("price",)
EMPTY_API_VALUE = "-"
VALUE_LABELS = {
    "total": "Brasil",
    "sc": "Santa Catarina",
}
STATE_NAMES = {
    "sc": "Santa Catarina",
}


@dataclass(frozen=True)
class Product:
    code: str
    name: str
    columns: dict[str, str]
    price_col: str | None  # None = produto sem coluna de preco na planilha
    price_kind: str  # "unit" or "kg"
    ncm_codes: tuple[str, ...] = ()


# A planilha mantem uma unica serie de preco: coluna W (US$/unid) para o
# produto 850423. Os demais produtos ficam com price_col=None.
PRODUCTS: tuple[Product, ...] = (
    Product("850421", "Transformadores dieletricos liquido <= 650 kVA", {"total": "F", "sc": "G"}, None, "unit", ("85042100",)),
    Product("850422", "Transformadores dieletricos liquido > 650 e <= 10.000 kVA", {"total": "H", "sc": "I"}, None, "unit", ("85042200",)),
    Product("850423", "Transformadores dieletricos liquido > 10.000 kVA", {"total": "J", "sc": "K"}, "W", "unit", ("85042300",)),
    Product("850431", "Outros transformadores eletricos <= 1 kVA", {"total": "L", "sc": "M"}, None, "kg"),
    Product("850432", "Outros transformadores eletricos > 1 e <= 16 kVA", {"total": "N", "sc": "O"}, None, "unit", ("85043229", "85043219", "85043221", "85043211")),
    Product("850433", "Outros transformadores eletricos > 16 e <= 500 kVA", {"total": "P", "sc": "Q"}, None, "unit", ("85043300",)),
    Product("850434", "Outros transformadores eletricos > 500 kVA", {"total": "R", "sc": "S"}, None, "unit", ("85043400",)),
    Product("850490", "Partes de transformadores", {"total": "T", "sc": "U"}, None, "kg"),
)


def log(message: str) -> None:
    print(f"[{datetime.now():%H:%M:%S}] {message}")


def parse_month(value) -> tuple[int, int] | None:
    try:
        if isinstance(value, (datetime, date)):
            return value.year, value.month
        if isinstance(value, str):
            text = value.strip()
            if not text:
                return None
            if "-" in text:
                parts = text.split("-")
                return int(parts[0]), int(parts[1])
            if "/" in text:
                parts = text.split("/")
                if len(parts[0]) == 4:
                    return int(parts[0]), int(parts[1])
                return int(parts[1]), int(parts[0])
    except (TypeError, ValueError, IndexError):
        pass
    return None


def has_value(value) -> bool:
    return value is not None and str(value).strip() != ""


def format_log_value(value) -> str:
    return f"{value:,.2f}" if isinstance(value, (int, float)) else str(value)


def build_payload(
    filter_name: str,
    values: list[str] | tuple[str, ...],
    year_start: int,
    year_end: int,
    metrics: list[str],
) -> dict:
    return {
        "flow": "export",
        "monthDetail": True,
        "period": {"from": f"{year_start}-01", "to": f"{year_end}-12"},
        "filters": [{"filter": filter_name, "values": list(values)}],
        "details": ["state", "subHeading"] if filter_name == "subHeading" else ["ncm"],
        "metrics": metrics,
    }


def comex_get(
    label: str,
    filter_name: str,
    values: list[str] | tuple[str, ...],
    year_start: int,
    year_end: int,
    metrics: list[str],
) -> list[dict]:
    payload = build_payload(filter_name, values, year_start, year_end, metrics)
    for attempt in range(1, MAX_RETRY + 1):
        try:
            response = requests.post(API_URL, json=payload, headers=API_HEADERS, timeout=90)
            if response.status_code == 429:
                log(f"Rate limit; aguardando {WAIT_429}s ({attempt}/{MAX_RETRY})...")
                time.sleep(WAIT_429)
                continue
            response.raise_for_status()
            body = response.json()
            return body.get("data", {}).get("list", []) or []
        except requests.RequestException as exc:
            log(f"Erro na API para {label} ({attempt}/{MAX_RETRY}): {exc}")
            if attempt < MAX_RETRY:
                time.sleep(WAIT_429)
    return []


def aggregate_values(rows: list[dict]) -> dict[tuple[int, int], dict[str, float]]:
    buckets: dict[tuple[int, int], dict[str, float]] = {}
    for row in rows:
        try:
            year = int(row.get("year") or row.get("coAno"))
            month = int(row.get("monthNumber") or row.get("coMes"))
            fob = float(row.get("metricFOB") or row.get("vlFob") or 0)
        except (TypeError, ValueError):
            continue

        bucket = buckets.setdefault((year, month), {"total": 0.0, "sc": 0.0})
        bucket["total"] += fob

        state = str(row.get("state") or row.get("noUf") or "").strip()
        for key, state_name in STATE_NAMES.items():
            if state == state_name:
                bucket[key] += fob

    return {month: {key: value / 1000 for key, value in values.items()} for month, values in buckets.items()}


def aggregate_price(rows: list[dict], denominator_field: str) -> dict[tuple[int, int], float]:
    buckets: dict[tuple[int, int], dict[str, float]] = {}
    for row in rows:
        try:
            year = int(row.get("year") or row.get("coAno"))
            month = int(row.get("monthNumber") or row.get("coMes"))
            fob = float(row.get("metricFOB") or row.get("vlFob") or 0)
            denominator = float(row.get(denominator_field) or 0)
        except (TypeError, ValueError):
            continue

        bucket = buckets.setdefault((year, month), {"fob": 0.0, "denominator": 0.0})
        bucket["fob"] += fob
        bucket["denominator"] += denominator

    result = {}
    for month, values in buckets.items():
        if values["denominator"]:
            result[month] = values["fob"] / values["denominator"]
    return result


def fetch_values(product: Product, first_year: int) -> dict[tuple[int, int], dict[str, float]]:
    current_year = datetime.today().year
    log(f"Consultando valor {product.code} | {first_year}-{current_year}...")
    rows = comex_get(product.code, "subHeading", [product.code], first_year, current_year, ["metricFOB"])
    data = aggregate_values(rows)
    log(f"  valor {product.code}: {len(rows)} linhas da API, {len(data)} meses.")
    return data


def fetch_prices(product: Product, first_year: int) -> dict[tuple[int, int], float]:
    current_year = datetime.today().year
    if product.price_kind == "unit":
        if not product.ncm_codes:
            log(f"  preco {product.code}: sem NCM mapeado; ignorando US$/unid.")
            return {}
        filter_name = "ncm"
        values = product.ncm_codes
        metrics = ["metricFOB", "metricStatistic"]
        denominator_field = "metricStatistic"
    else:
        filter_name = "subHeading"
        values = [product.code]
        metrics = ["metricFOB", "metricKG"]
        denominator_field = "metricKG"

    log(f"Consultando preco {product.code} ({product.price_kind}) | {first_year}-{current_year}...")
    rows = comex_get(product.code, filter_name, values, first_year, current_year, metrics)
    data = aggregate_price(rows, denominator_field)
    log(f"  preco {product.code}: {len(rows)} linhas da API, {len(data)} meses.")
    return data


def read_column(sheet, column: str, last_row: int) -> list:
    ci = column_index_from_string(column)
    return [sheet.cell(r, ci).value for r in range(FIRST_DATA_ROW, last_row + 1)]


def find_missing_cells(sheet, last_row: int):
    months = read_column(sheet, "B", last_row)
    current = {
        product.code: {
            **{key: read_column(sheet, column, last_row) for key, column in product.columns.items()},
            "price": read_column(sheet, product.price_col, last_row) if product.price_col else [],
        }
        for product in PRODUCTS
    }
    rows_by_month = {}
    missing = {product.code: {key: set() for key in MISSING_KEYS} for product in PRODUCTS}

    for index, month_value in enumerate(months):
        month = parse_month(month_value)
        if month is None:
            continue
        rows_by_month[month] = (FIRST_DATA_ROW + index, index)
        for product in PRODUCTS:
            for key in MISSING_KEYS:
                if key == "price" and not product.price_col:
                    continue  # produto sem coluna de preco na planilha
                if not has_value(current[product.code][key][index]):
                    missing[product.code][key].add(month)

    return rows_by_month, missing


# Coluna derivada, calculada na planilha e nao pela API: Volume (unid.) =
# total (mil US$) / preco unitario (US$/unid). Grava-se a FORMULA, nunca o valor
# resolvido -- assim a celula continua viva no Excel, como as vizinhas.
#
# Sem isto a coluna congela: o openpyxl escreve celula e nao propaga formula,
# entao a cada mes novo o coletor preenche J e W e X fica para tras, um mes por
# vez, ate alguem notar e arrastar na mao. O dashboard nao le X, entao a
# validacao nunca reclamaria.
VOLUME_COL = "X"
VOLUME_TOTAL_COL = "J"    # total do produto 850423
VOLUME_PRICE_COL = "W"    # preco unitario do mesmo produto
VOLUME_FORMULA = "={total}{row}*1000/{preco}{row}"


def estender_formula_volume(sheet, last_row: int, dry_run: bool) -> int:
    """Garante a formula de Volume em toda linha que ja tenha total e preco.

    Varre tudo em vez de so as linhas recem-escritas: assim tambem fecha buracos
    deixados por execucoes anteriores, sem depender de alguem lembrar.
    """
    escritas = 0
    for row in range(FIRST_DATA_ROW, last_row + 1):
        total = sheet[f"{VOLUME_TOTAL_COL}{row}"].value
        preco = sheet[f"{VOLUME_PRICE_COL}{row}"].value
        # "-" (EMPTY_API_VALUE) e vazio nao rendem divisao: a formula daria
        # #VALUE! e sujaria a planilha.
        if not isinstance(total, (int, float)) or not isinstance(preco, (int, float)):
            continue
        celula = sheet[f"{VOLUME_COL}{row}"]
        formula = VOLUME_FORMULA.format(
            total=VOLUME_TOTAL_COL, preco=VOLUME_PRICE_COL, row=row
        )
        if celula.value == formula:
            continue
        if not dry_run:
            celula.value = formula
        escritas += 1
    return escritas


def process_workbook(workbook, dry_run: bool) -> dict:
    if SHEET_NAME not in workbook.sheetnames:
        raise RuntimeError(f"Aba {SHEET_NAME!r} nao encontrada.")

    sheet = workbook[SHEET_NAME]
    # A aba tem max_row inflado (centenas de milhares de linhas vazias);
    # o fim real dos dados e' o fim do grid de datas da coluna B.
    last_row = FIRST_DATA_ROW
    for r in range(FIRST_DATA_ROW, sheet.max_row + 1):
        if sheet.cell(r, 2).value is None:
            break
        last_row = r
    rows_by_month, missing = find_missing_cells(sheet, last_row)

    volume_escritas = estender_formula_volume(sheet, last_row, dry_run)
    if volume_escritas:
        prefix = "[DRY-RUN] " if dry_run else ""
        log(f"{prefix}Coluna {VOLUME_COL} (Volume): {volume_escritas} formula(s) estendida(s).")

    products_to_fetch = [
        product
        for product in PRODUCTS
        if any(missing[product.code][key] for key in MISSING_KEYS)
    ]
    if not products_to_fetch:
        log("Nada a atualizar; colunas SECEX ja estao preenchidas.")
        return {"months": [], "written": volume_escritas}

    written = volume_escritas
    updated_months = set()

    for product_index, product in enumerate(products_to_fetch, start=1):
        missing_value_months = sorted(set().union(*(missing[product.code][key] for key in VALUE_KEYS)))
        missing_price_months = sorted(missing[product.code]["price"])

        log(f"{product.code} ({product_index}/{len(products_to_fetch)})")

        value_data = {}
        if missing_value_months:
            value_data = fetch_values(product, min(year for year, _ in missing_value_months))
        value_last_month = max(value_data) if value_data else None

        price_data = {}
        if missing_price_months:
            price_data = fetch_prices(product, min(year for year, _ in missing_price_months))
        price_last_month = max(price_data) if price_data else None

        for month in sorted(set(missing_value_months) | set(missing_price_months)):
            row_number, _ = rows_by_month[month]
            pending = []

            values = value_data.get(month)
            if values is not None:
                for key in VALUE_KEYS:
                    if month in missing[product.code][key]:
                        pending.append((product.columns[key], values[key], VALUE_LABELS[key]))
            else:
                if value_last_month is not None and month <= value_last_month:
                    for key in VALUE_KEYS:
                        if month in missing[product.code][key]:
                            pending.append((product.columns[key], EMPTY_API_VALUE, VALUE_LABELS[key]))

            if month in missing[product.code]["price"] and month in price_data:
                pending.append((product.price_col, price_data[month], "Preco"))
            elif (
                month in missing[product.code]["price"]
                and price_last_month is not None
                and month <= price_last_month
            ):
                pending.append((product.price_col, EMPTY_API_VALUE, "Preco"))

            if not pending:
                continue

            if not dry_run:
                for column, value, _ in pending:
                    sheet.cell(row_number, column_index_from_string(column)).value = value

            written += len(pending)
            updated_months.add(month)
            prefix = "[DRY-RUN] " if dry_run else ""
            labels = ", ".join(f"{label} {format_log_value(value)}" for _, value, label in pending)
            log(f"{prefix}{product.code} | {month[1]:02d}/{month[0]} | {labels}")

        if product_index < len(products_to_fetch):
            time.sleep(WAIT_NORMAL)

    return {"months": sorted(updated_months), "written": written}


def open_workbook(read_only: bool):
    """Baixa a planilha do Google Drive e abre com openpyxl."""
    file_id = drive_io.find_file_id(WORKBOOK_NAME)
    path = drive_io.download_to_temp(file_id)
    workbook = openpyxl.load_workbook(path, keep_vba=True, data_only=False)
    return file_id, path, workbook


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Atualiza valor e preco SECEX de transformadores na planilha da WEG."
    )
    parser.add_argument("--dry-run", action="store_true", help="Consulta e mostra as alteracoes sem gravar.")
    args = parser.parse_args()

    workbook = None
    try:
        file_id, path, workbook = open_workbook(args.dry_run)
        result = process_workbook(workbook, args.dry_run)
        if result["written"] and not args.dry_run:
            workbook.save(path)
            drive_io.update_file(file_id, path)   # sobe de volta pro Drive
            log("Arquivo salvo.")
        elif not result["written"]:
            log("Nenhuma celula alterada.")
    finally:
        if workbook is not None:
            workbook.close()

    mode = "DRY-RUN" if args.dry_run else "CONCLUIDO"
    print(f"\n[{mode}] {len(result['months'])} mes(es) atualizado(s), {result['written']} valor(es) preenchido(s).")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        raise SystemExit(130)
    except Exception as exc:
        log(f"ERRO: {exc}")
        raise SystemExit(1)
