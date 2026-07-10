"""
atualiza_intended_placements_frangous.py

Atualiza duas colunas da aba "Production" no FrangoUS.xlsm:

  Coluna E  — Pullet Chicks Hatched (US domestic)  ('000 heads)
              Fonte: USDA NASS QuickStats API
              CHICKENS, CHICKS, BROILER TYPE, PULLET, FOR DOMESTIC USE
                - PLACEMENTS, INTENDED, MEASURED IN HEAD

  Coluna F  — Cumulative Potential Placements  ('000 heads)
              Fórmula Excel: soma dos 9 meses na janela [M-15, M-7] de E.
              Só exibe valor quando todos os 9 meses de E estão preenchidos.
              Fórmula arrastada para todas as linhas de dados (incluindo futuras).

Pré-requisitos:
  pip install requests openpyxl
"""

from __future__ import annotations

import datetime as dt
import os
import sys
import warnings
from pathlib import Path
import copy

import openpyxl
import requests

sys.stdout.reconfigure(encoding="utf-8")
warnings.filterwarnings("ignore")


# ── Configuração ──────────────────────────────────────────────────────────────

DRIVE_ROOT = Path(
    r"G:\.shortcut-targets-by-id\11q1ngdqbySnDMCQSP4fiwPztT6ju2F4Z"
)
XLSX_NAME      = "FrangoUS.xlsm"
SHEET_NAME     = "Production"
FIRST_DATA_ROW = 3
COL_DATE       = 2   # B
COL_E          = 5   # E — Pullet Chicks Hatched (domestic)
COL_F          = 6   # F — Cumulative Potential Placements

# NASS_API_KEY vem do ambiente. Localmente, um .env (fora do git) ao lado
# deste script preenche o ambiente; na nuvem, use GitHub Secrets.
_env_file = Path(__file__).resolve().parent / ".env"
if _env_file.exists():
    for _line in _env_file.read_text(encoding="utf-8-sig").splitlines():
        _line = _line.strip()
        if _line and not _line.startswith("#") and "=" in _line:
            _k, _v = _line.split("=", 1)
            os.environ.setdefault(_k.strip(), _v.strip().strip('"').strip("'"))

NASS_API_KEY  = os.environ.get("NASS_API_KEY")
if not NASS_API_KEY:
    raise SystemExit("Defina NASS_API_KEY (variavel de ambiente ou .env).")
NASS_BASE_URL = "https://quickstats.nass.usda.gov/api/api_GET/"

NASS_PARAMS = {
    "source_desc":    "SURVEY",
    "sector_desc":    "ANIMALS & PRODUCTS",
    "commodity_desc": "CHICKENS",
    "short_desc":     "CHICKENS, CHICKS, BROILER TYPE, PULLET, FOR DOMESTIC USE"
                      " - PLACEMENTS, INTENDED, MEASURED IN HEAD",
    "agg_level_desc": "NATIONAL",
    "freq_desc":      "MONTHLY",
    "format":         "JSON",
}

MONTH_MAP = {
    "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6,
    "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12,
}


# ── Helpers ───────────────────────────────────────────────────────────────────

def localizar_planilha() -> Path:
    """Localiza a planilha sem depender dos nomes das pastas intermediárias."""
    encontrados = list(DRIVE_ROOT.rglob(XLSX_NAME))

    if not encontrados:
        raise FileNotFoundError(
            f"Planilha {XLSX_NAME!r} não encontrada dentro de {DRIVE_ROOT}"
        )

    if len(encontrados) > 1:
        caminhos = "\n".join(f"  - {path}" for path in encontrados)
        raise RuntimeError(
            f"Foram encontradas várias planilhas chamadas {XLSX_NAME!r}:\n"
            f"{caminhos}"
        )

    return encontrados[0]


def _val_is_empty(v) -> bool:
    if v is None:
        return True
    if isinstance(v, str) and v.strip().lower() in ("", "sem dados"):
        return True
    return False


def _approx_equal(a: float, b: float, tol: float = 1e-9) -> bool:
    return abs(a - b) <= tol * max(1.0, abs(a), abs(b))


# ── 1. Download NASS ──────────────────────────────────────────────────────────

def fetch_nass() -> list[dict]:
    params = {**NASS_PARAMS, "key": NASS_API_KEY}
    print(f"Consultando NASS QuickStats...")
    print(f"  {NASS_PARAMS['short_desc']}")
    resp = requests.get(NASS_BASE_URL, params=params, timeout=60)

    if resp.status_code == 200:
        data = resp.json()
        if "data" in data:
            print(f"  {len(data['data'])} registros recebidos.")
            return data["data"]
        raise RuntimeError(f"Resposta NASS inesperada: {data}")

    if resp.status_code == 413:
        print("  Muitos registros — adicionando filtro year__GE=2013 ...")
        params["year__GE"] = "2013"
        resp = requests.get(NASS_BASE_URL, params=params, timeout=60)
        resp.raise_for_status()
        data = resp.json()
        if "data" in data:
            print(f"  {len(data['data'])} registros recebidos.")
            return data["data"]

    resp.raise_for_status()
    raise RuntimeError("Resposta NASS inesperada.")


# ── 2. Parse ──────────────────────────────────────────────────────────────────

def parse_nass(records: list[dict]) -> dict[tuple[int, int], float]:
    """Retorna {(ano, mes): valor_em_mil_cabecas}. Unidade HEAD → ÷1000."""
    result: dict[tuple[int, int], float] = {}
    skipped = 0

    for rec in records:
        period  = str(rec.get("reference_period_desc", "")).strip().upper()
        year_s  = str(rec.get("year", "")).strip()
        value_s = str(rec.get("Value", "")).strip().replace(",", "")

        month = MONTH_MAP.get(period[:3])
        if month is None:
            skipped += 1
            continue

        try:
            year  = int(year_s)
            value = float(value_s)
        except (ValueError, TypeError):
            skipped += 1
            continue

        result[(year, month)] = value / 1000.0

    print(f"  Parse: {len(result)} meses válidos ({skipped} ignorados).")
    if result:
        mn, mx = min(result), max(result)
        print(f"  Intervalo: {mn[0]}-{mn[1]:02d}  →  {mx[0]}-{mx[1]:02d}")
    return result


# ── 3. Fórmula cumulativa ─────────────────────────────────────────────────────

def cumulative_formula(row: int) -> str:
    """
    Retorna a fórmula Excel para a célula F{row}.
    Soma E[row-15] : E[row-7] somente quando todos os 9 valores existem.
    """
    top = row - 15
    bot = row - 7
    return f"=IF(COUNT(E{top}:E{bot})=9,SUM(E{top}:E{bot}),\"-\")"


# ── 4. Atualizar planilha ─────────────────────────────────────────────────────

def _get_ref_fmt(ws, col: int, first_row: int):
    """Retorna a célula de referência de formatação para a coluna."""
    ref = None
    for row in range(first_row, ws.max_row + 1):
        cell = ws.cell(row=row, column=col)
        if not _val_is_empty(cell.value) and cell.number_format not in ("General", ""):
            ref = cell
    if ref is None:
        candidate = ws.cell(row=first_row - 1, column=col)
        if candidate.number_format not in ("General", ""):
            ref = candidate
    # Fallback: usa col E como referência de formato
    if ref is None:
        ref = _get_ref_fmt(ws, COL_E, first_row) if col != COL_E else None
    return ref


def _apply_fmt(cell, ref) -> None:
    if ref is None:
        return
    cell.font          = copy.copy(ref.font)
    cell.fill          = copy.copy(ref.fill)
    cell.border        = copy.copy(ref.border)
    cell.alignment     = copy.copy(ref.alignment)
    cell.number_format = ref.number_format


def update_column(
    ws,
    col: int,
    data: dict[tuple[int, int], float],
    label: str,
    ref_fmt,
) -> tuple[int, int, list[str]]:
    """Preenche/revisa uma coluna. Retorna (novos, revisoes, log_lines)."""
    novos, revisoes = 0, 0
    logs: list[str] = []

    for row in range(FIRST_DATA_ROW, ws.max_row + 1):
        date_val = ws.cell(row=row, column=COL_DATE).value
        if not isinstance(date_val, (dt.datetime, dt.date)):
            continue

        ym      = (date_val.year, date_val.month)
        new_val = data.get(ym)
        if new_val is None:
            continue

        cell    = ws.cell(row=row, column=col)
        old_val = cell.value

        if _val_is_empty(old_val):
            cell.value = new_val
            _apply_fmt(cell, ref_fmt)
            novos += 1
            logs.append(f"  [{ym[0]}-{ym[1]:02d}] {label}: (vazio) → {new_val:,.1f}")
        elif isinstance(old_val, (int, float)):
            if not _approx_equal(float(old_val), new_val):
                cell.value = new_val
                _apply_fmt(cell, ref_fmt)
                revisoes += 1
                logs.append(
                    f"  [{ym[0]}-{ym[1]:02d}] {label}: {float(old_val):,.1f}"
                    f" → {new_val:,.1f}  (revisão)"
                )
            elif cell.number_format in ("General", ""):
                _apply_fmt(cell, ref_fmt)

    return novos, revisoes, logs


def update_production(nass_data: dict[tuple[int, int], float]) -> None:
    path = localizar_planilha()

    wb = openpyxl.load_workbook(str(path), keep_vba=True, data_only=False)
    try:
        ws = wb[SHEET_NAME]

        ref_e = _get_ref_fmt(ws, COL_E, FIRST_DATA_ROW)
        ref_f = _get_ref_fmt(ws, COL_F, FIRST_DATA_ROW) or ref_e

        # ── Coluna E: valores da API ──────────────────────────────────────────
        nov_e, rev_e, log_e = update_column(ws, COL_E, nass_data, "E", ref_e)

        # ── Coluna F: fórmulas só até a última linha com resultado válido ────
        # Última linha com E numérico
        last_e_row = FIRST_DATA_ROW
        for row in range(FIRST_DATA_ROW, ws.max_row + 1):
            if isinstance(ws.cell(row=row, column=COL_E).value, (int, float)):
                last_e_row = row
        # COUNT=9 é satisfeito até row onde E[row-7] = last_e_row → row = last_e_row+7
        last_f_row = last_e_row + 7

        formulas_written = 0
        for row in range(FIRST_DATA_ROW, last_f_row + 1):
            date_val = ws.cell(row=row, column=COL_DATE).value
            if not isinstance(date_val, (dt.datetime, dt.date)):
                continue
            if row - 15 < 1:
                continue
            cell    = ws.cell(row=row, column=COL_F)
            formula = cumulative_formula(row)
            if cell.value != formula:
                cell.value = formula
                # Copia formato da célula imediatamente acima
                above = ws.cell(row=row - 1, column=COL_F)
                if above.number_format not in ("General", ""):
                    _apply_fmt(cell, above)
                elif ref_f:
                    _apply_fmt(cell, ref_f)
                formulas_written += 1

        # Limpa fórmulas que existam além de last_f_row (caso o range tenha encolhido)
        for row in range(last_f_row + 1, ws.max_row + 1):
            cell = ws.cell(row=row, column=COL_F)
            if cell.value is not None and str(cell.value).startswith("="):
                cell.value = None

        total = nov_e + rev_e + formulas_written
        if total > 0:
            wb.save(str(path))
            print(f"\nSalvo.")
            print(f"  Col E: {nov_e} nova(s), {rev_e} revisão/revisões")
            print(f"  Col F: {formulas_written} fórmula(s) escritas")
            for line in log_e[:60]:
                print(line)
            if len(log_e) > 60:
                print(f"  ... (+{len(log_e) - 60} mais)")
        else:
            print("Nenhuma alteração necessária — planilha já está atualizada.")
    finally:
        wb.close()


# ── 5. Main ───────────────────────────────────────────────────────────────────

def main() -> None:
    print(f"[{dt.datetime.now():%Y-%m-%d %H:%M}] atualiza_intended_placements_frangous.py")
    print(f"Arquivo : {XLSX_NAME}  |  Aba: {SHEET_NAME}")

    records   = fetch_nass()
    nass_data = parse_nass(records)

    if not nass_data:
        print("Nenhum dado mensal encontrado. Verifique os parâmetros NASS.")
        return

    print(f"\nAtualizando {XLSX_NAME} ...")
    update_production(nass_data)
    print(f"\n[{dt.datetime.now():%Y-%m-%d %H:%M}] Concluído.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        raise SystemExit(130)
