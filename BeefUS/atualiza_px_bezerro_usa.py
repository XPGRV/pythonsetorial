"""
Atualiza coluna N (Px Bezerro) da aba BeefUS com dados do USDA NASS.
Fonte: CATTLE, CALVES - PRICE RECEIVED, MEASURED IN $ / CWT (mensal, nacional)
"""

import copy
import requests
import openpyxl
from datetime import datetime
from pathlib import Path

# ── Configuração ──────────────────────────────────────────────────────────────
import os

# API_KEY (NASS) vem do ambiente. Localmente, um .env (fora do git) ao lado
# deste script preenche o ambiente; na nuvem, use GitHub Secrets.
_env_file = Path(__file__).resolve().parent / ".env"
if _env_file.exists():
    for _line in _env_file.read_text(encoding="utf-8-sig").splitlines():
        _line = _line.strip()
        if _line and not _line.startswith("#") and "=" in _line:
            _k, _v = _line.split("=", 1)
            os.environ.setdefault(_k.strip(), _v.strip().strip('"').strip("'"))

API_KEY      = os.environ.get("NASS_API_KEY")
if not API_KEY:
    raise SystemExit("Defina NASS_API_KEY (variavel de ambiente ou .env).")
DRIVE_ROOT   = Path(
    r"G:\.shortcut-targets-by-id\11q1ngdqbySnDMCQSP4fiwPztT6ju2F4Z"
)
XLSX_NAME    = "BeefUS.xlsm"
ABA          = "BeefUS"
COLUNA_DATA  = 2   # coluna B
COLUNA_PRECO = 14  # coluna N
LINHA_INICIO = 5   # primeira linha de dados (após cabeçalhos)
# ─────────────────────────────────────────────────────────────────────────────

MESES_EN = {
    "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4,
    "MAY": 5, "JUN": 6, "JUL": 7, "AUG": 8,
    "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12,
}


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


def buscar_precos_usda(api_key: str) -> dict[tuple[int, int], float]:
    """Retorna dict {(ano, mes): preco} com todos os dados disponíveis na API."""
    url = "https://quickstats.nass.usda.gov/api/api_GET/"
    params = {
        "key":                api_key,
        "source_desc":        "SURVEY",
        "sector_desc":        "ANIMALS & PRODUCTS",
        "group_desc":         "LIVESTOCK",
        "commodity_desc":     "CATTLE",
        "statisticcat_desc":  "PRICE RECEIVED",
        "short_desc":         "CATTLE, CALVES - PRICE RECEIVED, MEASURED IN $ / CWT",
        "domain_desc":        "TOTAL",
        "agg_level_desc":     "NATIONAL",
        "state_name":         "US TOTAL",
        "freq_desc":          "MONTHLY",
        "format":             "JSON",
    }

    print("Consultando USDA NASS API...")
    resp = requests.get(url, params=params, timeout=120)
    resp.raise_for_status()

    dados = resp.json().get("data", [])
    if not dados:
        raise ValueError("API não retornou dados. Verifique a chave ou os parâmetros.")

    precos = {}
    for item in dados:
        try:
            ano = int(item["year"])
            mes = MESES_EN[item["reference_period_desc"].strip().upper()]
            valor_raw = item["Value"].replace(",", "").strip()
            if valor_raw in ("", "(D)", "(NA)", "(Z)"):
                continue
            precos[(ano, mes)] = float(valor_raw)
        except (KeyError, ValueError):
            continue

    print(f"  {len(precos)} registros recebidos ({min(precos)} a {max(precos)})")
    return precos


def atualizar_planilha(precos: dict[tuple[int, int], float]) -> None:
    path = localizar_planilha()
    print(f"\nAbrindo planilha: {path}")
    wb = openpyxl.load_workbook(path, keep_vba=True)
    ws = wb[ABA]

    # captura formatação da última célula preenchida na coluna N
    ref_fmt = None
    for row in range(LINHA_INICIO, ws.max_row + 1):
        if ws.cell(row, COLUNA_PRECO).value is not None:
            ref_fmt = ws.cell(row, COLUNA_PRECO)

    atualizados = 0
    sem_preco   = 0

    for row in range(LINHA_INICIO, ws.max_row + 1):
        data_cell = ws.cell(row, COLUNA_DATA).value

        if data_cell is None:
            continue

        if isinstance(data_cell, datetime):
            chave = (data_cell.year, data_cell.month)
        else:
            try:
                chave = (data_cell.year, data_cell.month)
            except AttributeError:
                continue

        if chave in precos:
            cell = ws.cell(row, COLUNA_PRECO)
            nova_celula = cell.value is None
            cell.value = precos[chave]
            if nova_celula and ref_fmt is not None:
                cell.font          = copy.copy(ref_fmt.font)
                cell.fill          = copy.copy(ref_fmt.fill)
                cell.border        = copy.copy(ref_fmt.border)
                cell.alignment     = copy.copy(ref_fmt.alignment)
                cell.number_format = ref_fmt.number_format
            atualizados += 1
        else:
            sem_preco += 1

    print(f"  {atualizados} células atualizadas, {sem_preco} datas sem dado na API")
    print("Salvando...")
    wb.save(path)
    print("Concluído.")


if __name__ == "__main__":
    if API_KEY == "SUA_API_KEY_AQUI":
        print("ERRO: edite o script e insira sua API key do USDA NASS.")
        print("Solicite em: https://quickstats.nass.usda.gov/api")
    else:
        precos = buscar_precos_usda(API_KEY)
        atualizar_planilha(precos)
