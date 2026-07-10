"""Atualiza a aba Forecast do FrangoUS com Broilers do USDA ERS LDP."""

from __future__ import annotations

import datetime as dt
import re
import sys
import tempfile
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import drive_io
from BeefUS import coletar_usda_ers_ldp_forecasts as ers

WORKBOOK_NAME = "FrangoUS.xlsm"
SHEET_NAME = "Forecast"
ITEM_NAME = "Broilers"

EN_TO_PT = {
    "january": "jan", "february": "fev", "march": "mar", "april": "abr",
    "may": "mai", "june": "jun", "july": "jul", "august": "ago",
    "september": "set", "october": "out", "november": "nov", "december": "dez",
}


def release_matches_token(release: ers.Release, token_date: dt.date) -> bool:
    if release.date and (release.date.year, release.date.month) == (token_date.year, token_date.month):
        return True

    match = re.search(r":\s*([A-Za-z]+)\s+(\d{4})", release.title)
    if not match:
        return False
    month_name, year_text = match.groups()
    month_name = month_name.lower()
    return (
        int(year_text) == token_date.year
        and month_name in EN_TO_PT
        and list(EN_TO_PT).index(month_name) + 1 == token_date.month
    )


def main() -> None:
    print(f"Baixando {WORKBOOK_NAME} do Drive...")
    file_id = drive_io.find_file_id(WORKBOOK_NAME)
    workbook_path = Path(drive_io.download_to_temp(file_id))

    print(f"Verificando colunas vazias em {WORKBOOK_NAME} ({SHEET_NAME})...")
    missing = ers.find_empty_tokens(workbook_path, sheet_name=SHEET_NAME)
    if not missing:
        print(f"Nenhuma coluna vazia encontrada. {WORKBOOK_NAME} ja atualizado.")
        return

    print(f"FrangoUS colunas vazias: {', '.join(missing)}")
    today = dt.date.today()
    releases = ers.list_releases()
    production_rows: list[pd.DataFrame] = []

    with tempfile.TemporaryDirectory(prefix="ers-ldp-frangous-") as temp_dir:
        downloads_dir = Path(temp_dir)
        for token in missing:
            token_date = ers._token_to_approx_date(token)
            if not token_date or token_date > today:
                continue

            release = next((item for item in releases if release_matches_token(item, token_date)), None)
            if release is None:
                release = ers._discover_release_by_pubid_range(
                    ers._session(), token_date.year, token_date.month, dict(ers._KNOWN_PUBIDS)
                )
            if release is None:
                print(f"[WARN] {token}: release ERS nao encontrado.")
                continue

            downloads = ers.get_release_downloads(release)
            if not downloads.red_meat_xlsx_url:
                print(f"[WARN] {token}: XLSX red meat nao encontrado.")
                continue

            safe_name = re.sub(r"[^a-zA-Z0-9._-]+", "_", release.title).strip("_")
            xlsx_path = downloads_dir / f"{safe_name}.xlsx"
            print(f"Baixando ERS: {release.title}")
            ers.download_file(downloads.red_meat_xlsx_url, xlsx_path)

            try:
                production_rows.append(
                    ers.extract_production_quarters_with_forecast_flag(
                        xlsx_path, release=release, item=ITEM_NAME
                    )
                )
            except Exception as exc:
                print(f"[WARN] {token}: falha ao extrair Broilers: {exc}")

    if not production_rows:
        print("Nenhum dado ERS de Broilers foi extraido.")
        return

    values, flags = ers.build_release_matrix(production_rows)
    all_rows = pd.concat(production_rows, ignore_index=True)
    release_meta = (
        all_rows[["release_title", "release_date"]]
        .drop_duplicates()
        .assign(release_date=lambda frame: pd.to_datetime(frame["release_date"], errors="coerce"))
    )
    release_meta["release_col"] = release_meta["release_title"].str.replace(
        "Livestock, Dairy, and Poultry Outlook:", "", regex=False
    ).str.strip()

    ers.update_beefus_production_sheet(
        beefus_path=workbook_path,
        sheet_name=SHEET_NAME,
        values=values,
        flags=flags,
        release_meta=release_meta,
    )
    drive_io.update_file(file_id, str(workbook_path))
    print(f"OK: {WORKBOOK_NAME} salvo no Drive.")


if __name__ == "__main__":
    main()
