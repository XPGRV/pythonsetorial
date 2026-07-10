#!/usr/bin/env python3
"""
Atualiza os VOLUMES de soja e milho (Brasil e Mato Grosso) na aba SECEX da
Transportes.xlsm via Comex Stat.

Produtos (SH4):
  1201 - Soja, mesmo triturada
  1005 - Milho

Colunas preenchidas (volume, 1000 t):
  D - Soja       (Brasil, soma de todas as UFs)
  E - Soja - MT  (Mato Grosso)
  F - Milho      (Brasil, soma de todas as UFs)
  G - Milho - MT (Mato Grosso)

As colunas A e B ja existem na planilha e nao sao alteradas.
Nada e escrito apos a coluna G.

Estrategia: uma consulta por produto detalhando por UF (noUf). O total do
Brasil e a soma de todas as UFs (inclui "Nao Declarada"); o MT e a linha da
UF "Mato Grosso". kgLiquido / 1e6 = 1000 t.

Uso:
  python update_secex.py
  python update_secex.py --dry-run

Dependencias:
  pip install requests xlwings
  Requer Microsoft Excel instalado.
"""

from __future__ import annotations

import argparse
import sys
import time
from datetime import date, datetime
from pathlib import Path

import requests
import xlwings as xw


sys.stdout.reconfigure(encoding="utf-8")


# ---------------------------------------------------------------------------
# Configuracao
# ---------------------------------------------------------------------------
WORKBOOK_PATH = Path(
    r"G:\.shortcut-targets-by-id\11q1ngdqbySnDMCQSP4fiwPztT6ju2F4Z"
    r"\Arquivos\Setorial - Database\Transportes.xlsm"
)
SHEET_NAME = "SECEX"
FIRST_DATA_ROW = 3

API_URL = "https://api-comexstat.mdic.gov.br/general"
# Cabecalhos de navegador — a API fica atras do Cloudflare e recusa clientes
# "crus". Do IP residencial estes headers passam.
API_HEADERS = {"Content-Type": "application/json"}

WAIT_NORMAL = 13
WAIT_429 = 25
MAX_RETRY = 4

PRODUCTS = {
    "1201": "Soja",
    "1005": "Milho",
}
UF_MT = "Mato Grosso"  # noUf exato (distinto de "Mato Grosso do Sul")

COL_MONTH = "B"
COL_SOY = "D"
COL_SOY_MT = "E"
COL_CORN = "F"
COL_CORN_MT = "G"

TOLERANCE = 1e-3  # kt


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
# API Comex Stat (GET /general?filter=<json>) — detalhado por UF
# ---------------------------------------------------------------------------
def build_filter(code: str, year_start: int, year_end: int) -> dict:
    return {
        "yearStart": str(year_start),
        "yearEnd": str(year_end),
        "typeForm": 1,
        "typeOrder": 2,
        "filterList": [{
            "id": "noSh4pt", "text": "Posição (SH4)",
            "route": "/pt/harmonized-system/position", "type": "2",
            "group": "sh", "groupText": "Sistema Harmonizado (SH)",
            "hint": "fieldsForm.general.noSh4.description",
            "placeholder": "Posição (SH4)",
        }],
        "filterArray": [{"item": [code], "idInput": "noSh4pt"}],
        "rangeFilter": [],
        "detailDatabase": [
            {"id": "noSh4pt", "text": "Posição (SH4)", "parentId": "coSh4",
             "parent": "Código SH4", "group": "sh",
             "groupText": "Sistema Harmonizado (SH)"},
            {"id": "noUf", "text": "UF do Produto", "group": "gerais",
             "groupText": "Gerais"},
        ],
        "monthDetail": True,
        "metricFOB": False,
        "metricKG": True,
        "metricStatistic": False,
        "metricFreight": False,
        "metricInsurance": False,
        "metricCIF": False,
        "monthStart": "01",
        "monthEnd": "12",
        "formQueue": "general",
        "langDefault": "pt",
        "monthStartName": "Janeiro",
        "monthEndName": "Dezembro",
    }


def build_payload(code: str, year_start: int, year_end: int) -> dict:
    return {
        "flow": "export",
        "monthDetail": True,
        "period": {"from": f"{year_start}-01", "to": f"{year_end}-12"},
        "filters": [{"filter": "heading", "values": [code]}],
        "details": ["state", "heading"],
        "metrics": ["metricKG"],
    }


def comex_get(code: str, year_start: int, year_end: int) -> list[dict]:
    payload = build_payload(code, year_start, year_end)
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
            log(f"Erro na API ({attempt}/{MAX_RETRY}): {exc}")
            if attempt < MAX_RETRY:
                time.sleep(WAIT_429)
    return []


def aggregate(rows: list[dict]) -> dict:
    """(year, month) -> {'total': kt|None, 'mt': kt|None} a partir das linhas por UF."""
    buckets: dict[tuple[int, int], dict[str, float]] = {}
    for row in rows:
        try:
            year = int(row.get("year") or row.get("coAno"))
            month = int(row.get("monthNumber") or row.get("coMes"))
        except (KeyError, ValueError, TypeError):
            continue
        try:
            kg = float(row.get("metricKG") or row.get("kgLiquido") or 0)
        except (TypeError, ValueError):
            kg = 0.0
        bucket = buckets.setdefault((year, month), {"total_kg": 0.0, "mt_kg": 0.0})
        bucket["total_kg"] += kg
        if str(row.get("state") or row.get("noUf") or "").strip() == UF_MT:
            bucket["mt_kg"] += kg

    result = {}
    for month, bucket in buckets.items():
        result[month] = {
            "total": bucket["total_kg"] / 1e6,
            "mt": bucket["mt_kg"] / 1e6,
        }
    return result


def fetch_products(first_year: int) -> dict:
    current_year = datetime.today().year
    data = {}
    for code, name in PRODUCTS.items():
        log(f"Consultando {name} (SH4 {code}) por UF, {first_year}-{current_year}...")
        data[code] = aggregate(comex_get(code, first_year, current_year))
        time.sleep(WAIT_NORMAL)
    return data


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------
def process_workbook(workbook, dry_run: bool) -> dict:
    if SHEET_NAME not in [sheet.name for sheet in workbook.sheets]:
        raise RuntimeError(f"Aba {SHEET_NAME!r} nao encontrada.")

    sheet = workbook.sheets[SHEET_NAME]
    last_row = sheet.used_range.last_cell.row

    def col(column):
        values = sheet.range(f"{column}{FIRST_DATA_ROW}:{column}{last_row}").value
        return values if isinstance(values, list) else [values]

    month_values = col(COL_MONTH)
    current = {c: col(c) for c in (COL_SOY, COL_SOY_MT, COL_CORN, COL_CORN_MT)}

    rows_by_month = {}
    candidates = []
    for index, month_value in enumerate(month_values):
        month = parse_month(month_value)
        if month is None:
            continue
        row_number = FIRST_DATA_ROW + index
        rows_by_month[month] = (row_number, index)
        if any(not has_value(current[c][index]) for c in current):
            candidates.append(month)

    candidates = sorted(set(candidates))
    if not candidates:
        log("Nada a atualizar; volumes ja preenchidos.")
        return {"months": [], "written": 0}

    first_year = min(year for year, _ in candidates)
    log(
        f"Meses a preencher ({len(candidates)}): "
        f"{candidates[0][1]:02d}/{candidates[0][0]} a "
        f"{candidates[-1][1]:02d}/{candidates[-1][0]}"
    )
    data = fetch_products(first_year)
    soy = data.get("1201", {})
    corn = data.get("1005", {})

    updated_months = []
    written = 0

    for month in candidates:
        row_number, index = rows_by_month[month]
        s = soy.get(month)
        c = corn.get(month)
        if s is None and c is None:
            continue  # mes ainda nao disponivel na Comex Stat

        mapping = [
            (COL_SOY, s["total"] if s else None, "Soja"),
            (COL_SOY_MT, s["mt"] if s else None, "Soja-MT"),
            (COL_CORN, c["total"] if c else None, "Milho"),
            (COL_CORN_MT, c["mt"] if c else None, "Milho-MT"),
        ]

        changes = []
        for column, value, label in mapping:
            if value is None:
                continue
            cur = current[column][index]
            if not has_value(cur) or abs(float(cur) - value) > TOLERANCE:
                changes.append((column, value, label))

        if not changes:
            continue

        if not dry_run:
            for column, value, _ in changes:
                sheet.range(f"{column}{row_number}").value = value

        written += len(changes)
        updated_months.append(month)
        prefix = "[DRY-RUN] " if dry_run else ""
        log(
            f"{prefix}{month[1]:02d}/{month[0]} | "
            f"Soja {(s['total'] if s else 0):.1f} kt (MT {(s['mt'] if s else 0):.1f}) | "
            f"Milho {(c['total'] if c else 0):.1f} kt (MT {(c['mt'] if c else 0):.1f})"
        )

    return {"months": updated_months, "written": written}


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main() -> None:
    parser = argparse.ArgumentParser(
        description="Atualiza volumes de soja e milho (Brasil e MT) na aba SECEX."
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Consulta e mostra as alteracoes sem gravar a planilha.",
    )
    args = parser.parse_args()

    if not WORKBOOK_PATH.exists():
        raise FileNotFoundError(f"Planilha nao encontrada: {WORKBOOK_PATH}")

    app = xw.App(visible=False, add_book=False)
    try:
        app.api.AskToUpdateLinks = False
        app.display_alerts = False
        app.screen_updating = False
        workbook = app.books.open(
            str(WORKBOOK_PATH), update_links=0, read_only=args.dry_run
        )
        try:
            result = process_workbook(workbook, args.dry_run)
            if result["written"] and not args.dry_run:
                workbook.api.SaveLinkValues = True
                workbook.save()
                log("Arquivo salvo.")
            elif not result["written"]:
                log("Nenhuma celula alterada.")
        finally:
            workbook.close()
    finally:
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
