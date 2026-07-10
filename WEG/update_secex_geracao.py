#!/usr/bin/env python3
"""
Atualiza os dados SECEX de geracao na planilha WEG - Setorial.xlsm.

Fonte: Comex Stat API
Fluxo: exportacao
Detalhe: mes + UF do produto
Metrica: US$ FOB, gravado na planilha em 1000 US$

Planilha:
  G:\\.shortcut-targets-by-id\\11q1ngdqbySnDMCQSP4fiwPztT6ju2F4Z\\Arquivos\\Setorial - Database\\WEG - Setorial.xlsm
  Aba: Geracao

Colunas preenchidas em 1000 US$:
  D/E   - 850161 | Total Brasil / Santa Catarina
  F/G   - 850162 | Total Brasil / Santa Catarina
  H/I   - 850163 | Total Brasil / Santa Catarina
  J/K   - 850164 | Total Brasil / Santa Catarina
  L/M   - 850300 | Total Brasil / Santa Catarina
  N/O   - 850231 | Total Brasil / Santa Catarina
  P/Q   - 841011 | Total Brasil / Santa Catarina
  R/S   - 841012 | Total Brasil / Santa Catarina
  T/U   - 841013 | Total Brasil / Santa Catarina
  V/W   - 841090 | Total Brasil / Santa Catarina
  X/Y   - 840681 | Total Brasil / Santa Catarina
  Z/AA  - 840682 | Total Brasil / Santa Catarina
  AB/AC - 840690 | Total Brasil / Santa Catarina

Comportamento:
  - Escreve apenas celulas vazias.
  - Nao revisa nem sobrescreve valores ja preenchidos.
  - Consulta a Comex Stat apenas para produtos com celulas vazias.
  - Para cada produto, consulta a partir do primeiro ano vazio daquele produto.

Uso:
  python update_secex_geracao.py
  python update_secex_geracao.py --dry-run

Dependencias:
  pip install requests xlwings
  Requer Microsoft Excel instalado.
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


# ---------------------------------------------------------------------------
# Configuracao
# ---------------------------------------------------------------------------
WORKBOOK_PATH = Path(
    r"G:\.shortcut-targets-by-id\11q1ngdqbySnDMCQSP4fiwPztT6ju2F4Z\Arquivos\Setorial - Database\WEG - Setorial.xlsm"
)
SHEET_NAME = "Gera\u00e7\u00e3o"
FIRST_DATA_ROW = 5

API_URL = "https://api-comexstat.mdic.gov.br/general"
API_HEADERS = {"Content-Type": "application/json"}

WAIT_NORMAL = 8
WAIT_429 = 25
MAX_RETRY = 4

UF_SC = "Santa Catarina"


@dataclass(frozen=True)
class Product:
    code: str
    name: str
    total_col: str
    sc_col: str


PRODUCTS: tuple[Product, ...] = (
    Product("850161", "Geradores AC <= 75 kVA", "D", "E"),
    Product("850162", "Geradores AC > 75 kVA e <= 375 kVA", "F", "G"),
    Product("850163", "Geradores AC > 375 kVA e <= 750 kVA", "H", "I"),
    Product("850164", "Geradores AC > 750 kVA", "J", "K"),
    Product("850300", "Partes de motores, geradores e grupos eletrogenos", "L", "M"),
    Product("850231", "Grupos eletrogenos de energia eolica", "N", "O"),
    Product("841011", "Turbinas hidraulicas <= 1.000 kW", "P", "Q"),
    Product("841012", "Turbinas hidraulicas > 1.000 kW e <= 10.000 kW", "R", "S"),
    Product("841013", "Turbinas hidraulicas > 10.000 kW", "T", "U"),
    Product("841090", "Partes de turbinas hidraulicas", "V", "W"),
    Product("840681", "Turbinas a vapor > 40 MW", "X", "Y"),
    Product("840682", "Turbinas a vapor <= 40 MW", "Z", "AA"),
    Product("840690", "Partes de turbinas a vapor", "AB", "AC"),
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
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


# ---------------------------------------------------------------------------
# API Comex Stat
# ---------------------------------------------------------------------------
def build_payload(code: str, year_start: int, year_end: int) -> dict:
    return {
        "flow": "export",
        "monthDetail": True,
        "period": {"from": f"{year_start}-01", "to": f"{year_end}-12"},
        "filters": [{"filter": "subHeading", "values": [code]}],
        "details": ["state", "subHeading"],
        "metrics": ["metricFOB"],
    }


def comex_get(product: Product, year_start: int, year_end: int) -> list[dict]:
    payload = build_payload(product.code, year_start, year_end)
    for attempt in range(1, MAX_RETRY + 1):
        try:
            response = requests.post(
                API_URL,
                json=payload,
                headers=API_HEADERS,
                timeout=90,
            )
            if response.status_code == 429:
                log(f"Rate limit; aguardando {WAIT_429}s ({attempt}/{MAX_RETRY})...")
                time.sleep(WAIT_429)
                continue
            response.raise_for_status()
            body = response.json()
            return body.get("data", {}).get("list", []) or []
        except requests.RequestException as exc:
            log(f"Erro na API para {product.code} ({attempt}/{MAX_RETRY}): {exc}")
            if attempt < MAX_RETRY:
                time.sleep(WAIT_429)
    return []


def aggregate_usd_thousand(rows: list[dict]) -> dict[tuple[int, int], dict[str, float]]:
    """Retorna {(ano, mes): {'total': 1000 US$, 'sc': 1000 US$}}."""
    buckets: dict[tuple[int, int], dict[str, float]] = {}
    for row in rows:
        try:
            year = int(row.get("year") or row.get("coAno"))
            month = int(row.get("monthNumber") or row.get("coMes"))
        except (TypeError, ValueError):
            continue

        try:
            fob = float(row.get("metricFOB") or row.get("vlFob") or 0)
        except (TypeError, ValueError):
            fob = 0.0

        bucket = buckets.setdefault((year, month), {"total": 0.0, "sc": 0.0})
        bucket["total"] += fob
        if str(row.get("state") or row.get("noUf") or "").strip() == UF_SC:
            bucket["sc"] += fob

    return {
        month: {
            "total": values["total"] / 1000,
            "sc": values["sc"] / 1000,
        }
        for month, values in buckets.items()
    }


def fetch_product(product: Product, first_year: int) -> dict[tuple[int, int], dict[str, float]]:
    current_year = datetime.today().year
    log(f"Consultando {product.code} | {first_year}-{current_year}...")
    rows = comex_get(product, first_year, current_year)
    data = aggregate_usd_thousand(rows)
    log(f"  {product.code}: {len(rows)} linhas da API, {len(data)} meses.")
    return data


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------
def read_column(sheet, column: str, last_row: int) -> list:
    values = sheet.range(f"{column}{FIRST_DATA_ROW}:{column}{last_row}").value
    return values if isinstance(values, list) else [values]


def find_missing_cells(
    sheet,
    last_row: int,
) -> tuple[
    dict[tuple[int, int], tuple[int, int]],
    dict[str, dict[str, set[tuple[int, int]]]],
]:
    months = read_column(sheet, "B", last_row)
    current = {
        product.code: {
            "total": read_column(sheet, product.total_col, last_row),
            "sc": read_column(sheet, product.sc_col, last_row),
        }
        for product in PRODUCTS
    }

    rows_by_month = {}
    missing = {
        product.code: {"total": set(), "sc": set()}
        for product in PRODUCTS
    }

    for index, month_value in enumerate(months):
        month = parse_month(month_value)
        if month is None:
            continue
        row_number = FIRST_DATA_ROW + index
        rows_by_month[month] = (row_number, index)
        for product in PRODUCTS:
            if not has_value(current[product.code]["total"][index]):
                missing[product.code]["total"].add(month)
            if not has_value(current[product.code]["sc"][index]):
                missing[product.code]["sc"].add(month)

    return rows_by_month, missing


def process_workbook(workbook, dry_run: bool) -> dict:
    sheet_names = [sheet.name for sheet in workbook.sheets]
    if SHEET_NAME not in sheet_names:
        raise RuntimeError(f"Aba {SHEET_NAME!r} nao encontrada.")

    sheet = workbook.sheets[SHEET_NAME]
    last_row = sheet.used_range.last_cell.row

    rows_by_month, missing = find_missing_cells(sheet, last_row)
    products_to_fetch = [
        product
        for product in PRODUCTS
        if missing[product.code]["total"] or missing[product.code]["sc"]
    ]

    if not products_to_fetch:
        log("Nada a atualizar; colunas SECEX ja estao preenchidas.")
        return {"months": [], "written": 0}

    written = 0
    updated_months = set()

    for product_index, product in enumerate(products_to_fetch, start=1):
        missing_total = missing[product.code]["total"]
        missing_sc = missing[product.code]["sc"]
        missing_months = sorted(missing_total | missing_sc)
        first_year = min(year for year, _ in missing_months)
        log(
            f"{product.code} ({product_index}/{len(products_to_fetch)}) | "
            f"{len(missing_months)} mes(es) com vazio(s), "
            f"desde {missing_months[0][1]:02d}/{missing_months[0][0]}"
        )

        product_data = fetch_product(product, first_year)

        for month in missing_months:
            row_number, _ = rows_by_month[month]
            values = product_data.get(month)
            if values is None:
                continue

            pending = []
            if month in missing_total:
                pending.append((product.total_col, values["total"], "Brasil"))
            if month in missing_sc:
                pending.append((product.sc_col, values["sc"], "Santa Catarina"))
            if not pending:
                continue

            if not dry_run:
                for column, value, _ in pending:
                    sheet.range(f"{column}{row_number}").value = value

            written += len(pending)
            updated_months.add(month)
            prefix = "[DRY-RUN] " if dry_run else ""
            labels = ", ".join(f"{label} {value:,.1f}" for _, value, label in pending)
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


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main() -> None:
    parser = argparse.ArgumentParser(
        description="Atualiza 1000 US$ FOB SECEX de geracao na planilha da WEG."
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Consulta e mostra as alteracoes sem gravar a planilha.",
    )
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
    print(
        f"\n[{mode}] {len(result['months'])} mes(es) atualizado(s), "
        f"{result['written']} valor(es) preenchido(s)."
    )


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        raise SystemExit(130)
    except Exception as exc:
        log(f"ERRO: {exc}")
        raise SystemExit(1)
