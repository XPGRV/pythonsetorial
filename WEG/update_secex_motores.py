#!/usr/bin/env python3
"""
Atualiza os dados SECEX de motores e automacao na planilha WEG - Setorial.xlsm.

Fonte: Comex Stat API
Fluxo: exportacao
Detalhe: mes + UF do produto
Valor: US$ FOB, gravado em 1000 US$
Preco: US$/unid ou US$/kg, sem abertura por estado

Planilha:
  G:\\.shortcut-targets-by-id\\11q1ngdqbySnDMCQSP4fiwPztT6ju2F4Z\\Arquivos\\Setorial - Database\\WEG - Setorial.xlsm
  Aba: Motores

Comportamento:
  - Escreve apenas celulas vazias.
  - Nao revisa nem sobrescreve valores ja preenchidos.
  - Consulta a Comex Stat apenas para produtos/campos com celulas vazias.
  - Para cada produto, consulta a partir do primeiro ano vazio daquele produto.

Uso:
  python update_secex_motores.py
  python update_secex_motores.py --dry-run
"""

from __future__ import annotations

import argparse
import sys
import time
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

import requests
import xlwings as xw


sys.stdout.reconfigure(encoding="utf-8")


WORKBOOK_PATH = Path(
    r"G:\.shortcut-targets-by-id\11q1ngdqbySnDMCQSP4fiwPztT6ju2F4Z\Arquivos\Setorial - Database\WEG - Setorial.xlsm"
)
SHEET_NAME = "Motores"
FIRST_DATA_ROW = 5

API_URL = "https://api-comexstat.mdic.gov.br/general"
API_HEADERS = {"Content-Type": "application/json"}

WAIT_NORMAL = 8
WAIT_429 = 25
MAX_RETRY = 4

UF_SC = "Santa Catarina"
MISSING_KEYS = ("total", "sc", "price")
EMPTY_API_VALUE = "-"


@dataclass(frozen=True)
class Product:
    code: str
    name: str
    total_col: str
    sc_col: str
    price_col: str
    price_kind: str  # "unit" or "kg"
    ncm_codes: tuple[str, ...] = ()


PRODUCTS: tuple[Product, ...] = (
    Product("850120", "Motores eletricos universais de potencia > 37,5 W", "D", "E", "AR", "unit", ("85012000",)),
    Product("850131", "Motores e geradores CC de potencia <= 750 W", "F", "G", "AS", "unit", ("85013110", "85013120")),
    Product("850132", "Motores e geradores CC de potencia > 750 W e <= 75 kW", "H", "I", "AT", "unit", ("85013210", "85013220")),
    Product("850133", "Motores e geradores CC de potencia > 75 kW e <= 375 kW", "J", "K", "AU", "unit", ("85013310", "85013320")),
    Product("850134", "Motores e geradores CC de potencia > 375 kW", "L", "M", "AV", "unit", ("85013419", "85013420")),
    Product("850140", "Outros motores CA monofasicos", "N", "O", "AW", "unit", ("85014019", "85014011", "85014029", "85014021")),
    Product("850151", "Outros motores CA polifasicos <= 750 W", "P", "Q", "AX", "unit", ("85015110", "85015190", "85015120")),
    Product("850152", "Outros motores CA polifasicos > 750 W e <= 75 kW", "R", "S", "AY", "unit", ("85015210", "85015290", "85015220")),
    Product("850153", "Outros motores CA polifasicos > 75 kW", "T", "U", "AZ", "unit", ("85015310", "85015320", "85015390", "85015330")),
    Product("850440", "Conversores eletricos estaticos", "W", "X", "BA", "kg"),
    Product("850450", "Outras bobinas de reatancia e auto-inducao", "Y", "Z", "BB", "kg"),
    Product("853620", "Disjuntores para tensao <= 1 kV", "AA", "AB", "BC", "unit", ("85362000",)),
    Product("853521", "Disjuntores para tensao > 1 kV e < 72,5 kV", "AC", "AD", "BD", "unit", ("85352100",)),
    Product("853641", "Reles para tensao <= 60 V", "AE", "AF", "BE", "unit", ("85364100",)),
    Product("853649", "Outros reles para tensao > 60 V e <= 1.000 V", "AG", "AH", "BF", "unit", ("85364900",)),
    Product("853650", "Outros interruptores, seccionadores e comutadores <= 1 kV", "AI", "AJ", "BG", "kg"),
    Product("853690", "Outros aparelhos para interrupcao, seccionamento, protecao e ligacao <= 1 kV", "AK", "AL", "BH", "kg"),
    Product("853710", "Quadros e paineis para tensao <= 1 kV", "AM", "AN", "BI", "kg"),
    Product("853720", "Quadros e paineis para tensao > 1 kV", "AO", "AP", "BJ", "unit", ("85372090",)),
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
        if str(row.get("state") or row.get("noUf") or "").strip() == UF_SC:
            bucket["sc"] += fob

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
    values = sheet.range(f"{column}{FIRST_DATA_ROW}:{column}{last_row}").value
    return values if isinstance(values, list) else [values]


def find_missing_cells(sheet, last_row: int):
    months = read_column(sheet, "B", last_row)
    current = {
        product.code: {
            "total": read_column(sheet, product.total_col, last_row),
            "sc": read_column(sheet, product.sc_col, last_row),
            "price": read_column(sheet, product.price_col, last_row),
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
                if not has_value(current[product.code][key][index]):
                    missing[product.code][key].add(month)

    return rows_by_month, missing


def process_workbook(workbook, dry_run: bool) -> dict:
    if SHEET_NAME not in [sheet.name for sheet in workbook.sheets]:
        raise RuntimeError(f"Aba {SHEET_NAME!r} nao encontrada.")

    sheet = workbook.sheets[SHEET_NAME]
    last_row = sheet.used_range.last_cell.row
    rows_by_month, missing = find_missing_cells(sheet, last_row)

    products_to_fetch = [
        product
        for product in PRODUCTS
        if any(missing[product.code][key] for key in MISSING_KEYS)
    ]
    if not products_to_fetch:
        log("Nada a atualizar; colunas SECEX ja estao preenchidas.")
        return {"months": [], "written": 0}

    written = 0
    updated_months = set()

    for product_index, product in enumerate(products_to_fetch, start=1):
        missing_value_months = sorted(missing[product.code]["total"] | missing[product.code]["sc"])
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
                if month in missing[product.code]["total"]:
                    pending.append((product.total_col, values["total"], "Brasil"))
                if month in missing[product.code]["sc"]:
                    pending.append((product.sc_col, values["sc"], "Santa Catarina"))
            else:
                if value_last_month is not None and month <= value_last_month:
                    if month in missing[product.code]["total"]:
                        pending.append((product.total_col, EMPTY_API_VALUE, "Brasil"))
                    if month in missing[product.code]["sc"]:
                        pending.append((product.sc_col, EMPTY_API_VALUE, "Santa Catarina"))

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
                    sheet.range(f"{column}{row_number}").value = value

            written += len(pending)
            updated_months.add(month)
            prefix = "[DRY-RUN] " if dry_run else ""
            labels = ", ".join(f"{label} {format_log_value(value)}" for _, value, label in pending)
            log(f"{prefix}{product.code} | {month[1]:02d}/{month[0]} | {labels}")

        if product_index < len(products_to_fetch):
            time.sleep(WAIT_NORMAL)

    return {"months": sorted(updated_months), "written": written}


def open_workbook(read_only: bool):
    target = str(WORKBOOK_PATH).lower()
    for app in xw.apps:
        for book in app.books:
            try:
                if str(Path(book.fullname)).lower() == target:
                    log("Usando a planilha que ja esta aberta no Excel.")
                    return app, book, False
            except Exception:
                continue

    app = xw.App(visible=False, add_book=False)
    app.api.AskToUpdateLinks = False
    app.display_alerts = False
    app.screen_updating = False
    workbook = app.books.open(str(WORKBOOK_PATH), update_links=0, read_only=read_only)
    return app, workbook, True


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Atualiza valor e preco SECEX de motores na planilha da WEG."
    )
    parser.add_argument("--dry-run", action="store_true", help="Consulta e mostra as alteracoes sem gravar.")
    args = parser.parse_args()

    if not WORKBOOK_PATH.exists():
        raise FileNotFoundError(f"Planilha nao encontrada: {WORKBOOK_PATH}")

    app = None
    workbook = None
    opened_by_script = False
    try:
        app, workbook, opened_by_script = open_workbook(args.dry_run)
        result = process_workbook(workbook, args.dry_run)
        if result["written"] and not args.dry_run:
            workbook.api.SaveLinkValues = True
            workbook.save()
            log("Arquivo salvo.")
        elif not result["written"]:
            log("Nenhuma celula alterada.")
    finally:
        if workbook is not None and opened_by_script:
            workbook.close()
        if app is not None and opened_by_script:
            app.quit()

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
