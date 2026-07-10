#!/usr/bin/env python3
"""
Atualizacao de exportacoes de carnes — aba SECEX do BeefBR.xlsm
Execucao: python update_secex.py

Dependencias:
    pip install xlwings requests
    (requer Microsoft Excel instalado)
"""

import re
import sys
import requests
import time
from datetime import datetime, date
from pathlib import Path
import xlwings as xw

# ==============================================================================
# CONFIGURACOES
# ==============================================================================
DRIVE_ROOT = Path(
    r"G:\.shortcut-targets-by-id\11q1ngdqbySnDMCQSP4fiwPztT6ju2F4Z"
)
XLSX_NAME  = "BeefBR.xlsm"
SHEET_NAME = "SECEX"

COMEX_URL     = "https://api-comexstat.mdic.gov.br/general"
COMEX_HEADERS = {"Content-Type": "application/json"}

WAIT_NORMAL = 13
WAIT_429    = 25
MAX_RETRY   = 3


# ==============================================================================
# HELPERS
# ==============================================================================
def localizar_planilha() -> Path:
    """Localiza a planilha sem depender dos nomes das pastas intermediárias."""
    encontrados = list(DRIVE_ROOT.rglob(XLSX_NAME))

    if not encontrados:
        raise FileNotFoundError(
            f"Planilha {XLSX_NAME!r} nao encontrada dentro de {DRIVE_ROOT}"
        )

    if len(encontrados) > 1:
        caminhos = "\n".join(f"  - {path}" for path in encontrados)
        raise RuntimeError(
            f"Foram encontradas varias planilhas chamadas {XLSX_NAME!r}:\n"
            f"{caminhos}"
        )

    return encontrados[0]


def log(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")

def ym(y, m):
    return (int(y), int(m))

def parse_date(val):
    try:
        if isinstance(val, (datetime, date)):
            return val.year, val.month
        elif isinstance(val, str):
            val = val.strip()
            if "-" in val:
                p = val.split("-"); return int(p[0]), int(p[1])
            elif "/" in val:
                p = val.split("/")
                if len(p[0]) == 4: return int(p[0]), int(p[1])
                else:              return int(p[1]), int(p[0])
        return None
    except Exception:
        return None

def adjust_formula_row(formula, old_row, new_row):
    """Arrasta formula uma linha para baixo, respeitando referencias fixas ($)."""
    def replace_ref(m):
        col_part  = m.group(1)
        row_fixed = m.group(2)
        row_num   = int(m.group(3))
        if row_fixed == "" and row_num == old_row:
            return f"{col_part}{new_row}"
        return m.group(0)
    return re.sub(r"(\$?[A-Za-z]+)(\$?)(\d+)", replace_ref, formula)


# ==============================================================================
# API COMEX STAT
# ==============================================================================
def comex_post(payload, retry=MAX_RETRY):
    for attempt in range(retry):
        try:
            resp = requests.post(COMEX_URL, json=payload,
                                 headers=COMEX_HEADERS, timeout=60)
            if resp.status_code == 429:
                log(f"  Rate limit — aguardando {WAIT_429}s... (tentativa {attempt+1}/{retry})")
                time.sleep(WAIT_429)
                continue
            resp.raise_for_status()
            return resp.json()
        except requests.RequestException as e:
            log(f"  Erro na API (tentativa {attempt+1}/{retry}): {e}")
            time.sleep(WAIT_429)
    return None


# ==============================================================================
# MAIN
# ==============================================================================
def main():
    log("Verificando ultimo mes disponivel no Comex Stat...")
    hoje = datetime.today()
    result = comex_post({
        "flow": "export", "monthDetail": True,
        "period": {"from": "2024-01", "to": f"{hoje.year}-{hoje.month:02d}"},
        "filters": [{"filter": "heading", "values": ["0203"]}],
        "details": ["heading"], "metrics": ["metricFOB", "metricKG"],
    })
    time.sleep(WAIT_NORMAL)

    if not result or not result.get("data", {}).get("list"):
        log("ERRO: sem resposta do Comex Stat.")
        sys.exit(1)

    ultimo_mes_api = max(ym(r["year"], r["monthNumber"])
                         for r in result["data"]["list"])
    log(f"Ultimo mes disponivel: {ultimo_mes_api[1]:02d}/{ultimo_mes_api[0]}")

    try:
        path = localizar_planilha()
    except (FileNotFoundError, RuntimeError) as exc:
        log(f"ERRO: {exc}")
        sys.exit(1)

    app = xw.App(visible=False, add_book=False)
    try:
        app.api.AskToUpdateLinks = False
        app.display_alerts = False
        app.screen_updating = False

        wb = app.books.open(str(path), update_links=0, read_only=False)
        try:
            meses_ok, dados_final = processar(wb, ultimo_mes_api)
            if meses_ok:
                wb.api.SaveLinkValues = True
                wb.save()
                log("Arquivo salvo.")
            else:
                log("Nada alterado — planilha ja esta atualizada.")
        finally:
            wb.close()
    finally:
        app.quit()

    # Relatorio final
    if meses_ok:
        print("\n" + "=" * 60)
        print(f"RELATORIO — {len(meses_ok)} mes(es) atualizado(s)")
        print("=" * 60)
        for key in meses_ok:
            d  = dados_final[key]
            sk = d["suina_kt"]  or 0;  su = d["suina_usd"]  or 0
            bk = d["bovina_kt"] or 0;  bu = d["bovina_usd"] or 0
            fk = d["frango_kt"] or 0;  fu = d["frango_usd"] or 0
            print(f"\n  {key[1]:02d}/{key[0]}:")
            print(f"    Suina  : {sk:.3f} kt  |  {su:.4f} USD/kg")
            print(f"    Bovina : {bk:.3f} kt  |  {bu:.4f} USD/kg")
            print(f"    Frango : {fk:.3f} kt  |  {fu:.4f} USD/kg")
        print(f"\nPlanilha atualizada ate {ultimo_mes_api[1]:02d}/{ultimo_mes_api[0]}")


# ==============================================================================
# PROCESSAMENTO
# ==============================================================================
def processar(wb, ultimo_mes_api):
    sh = wb.sheets[SHEET_NAME]
    last_row = sh.used_range.last_cell.row

    dates_vals = sh.range(f"B2:B{last_row}").value
    d_vals     = sh.range(f"D2:D{last_row}").value

    if not isinstance(dates_vals, list): dates_vals = [dates_vals]
    if not isinstance(d_vals,     list): d_vals     = [d_vals]

    date_rows    = {}
    filled       = {}
    last_ref_row = None

    for i, (date_val, d_val) in enumerate(zip(dates_vals, d_vals)):
        parsed = parse_date(date_val)
        if parsed is None:
            continue
        yr, mo = parsed
        key = ym(yr, mo)
        rn  = i + 2
        date_rows[key] = rn
        is_filled = d_val is not None and str(d_val).strip() != ""
        filled[key] = is_filled
        if is_filled:
            last_ref_row = rn

    gaps = []
    ultimo_mes_planilha = None
    for key in sorted(date_rows):
        if key <= ultimo_mes_api:
            if filled[key]: ultimo_mes_planilha = key
            else:           gaps.append(key)

    meses_novos = []
    if ultimo_mes_planilha:
        yr, mo = ultimo_mes_planilha
        while True:
            mo += 1
            if mo > 12: mo = 1; yr += 1
            k = ym(yr, mo)
            if k > ultimo_mes_api: break
            if k in date_rows: meses_novos.append(k)

    meses_para_buscar = sorted(set(gaps + meses_novos))

    if not meses_para_buscar:
        return [], {}

    log(f"Meses a preencher ({len(meses_para_buscar)}): "
        f"{[f'{k[1]:02d}/{k[0]}' for k in meses_para_buscar]}")
    log("Consultando Comex Stat...")

    PRODUTOS = {
        "0203": "Suina",
        "0201": "Bovina_fresca",
        "0202": "Bovina_congelada",
        "0207": "Frango",
    }
    dados_brutos = {k: {} for k in meses_para_buscar}

    hoje = datetime.today()
    ano_min     = min(k[0] for k in meses_para_buscar)
    period_from = f"{ano_min}-01"
    period_to   = f"{hoje.year}-12"

    for cod, nome in PRODUTOS.items():
        log(f"  {nome}...")
        r = comex_post({
            "flow": "export", "monthDetail": True,
            "period": {"from": period_from, "to": period_to},
            "filters": [{"filter": "heading", "values": [cod]}],
            "details": ["heading"], "metrics": ["metricFOB", "metricKG"],
        })
        time.sleep(WAIT_NORMAL)
        if r and r.get("data", {}).get("list"):
            for rec in r["data"]["list"]:
                k = ym(rec["year"], rec["monthNumber"])
                if k in dados_brutos:
                    dados_brutos[k][cod] = {
                        "kg":  float(rec.get("metricKG",  0) or 0),
                        "fob": float(rec.get("metricFOB", 0) or 0),
                    }

    dados_final = {}
    for key in meses_para_buscar:
        d  = dados_brutos[key]
        sk = d.get("0203", {}).get("kg",  0);  sf = d.get("0203", {}).get("fob", 0)
        bk = d.get("0201", {}).get("kg",  0) + d.get("0202", {}).get("kg",  0)
        bf = d.get("0201", {}).get("fob", 0) + d.get("0202", {}).get("fob", 0)
        fk = d.get("0207", {}).get("kg",  0);  ff = d.get("0207", {}).get("fob", 0)
        dados_final[key] = {
            "suina_kt":   sk / 1e6 if sk else None,
            "bovina_kt":  bk / 1e6 if bk else None,
            "frango_kt":  fk / 1e6 if fk else None,
            "suina_usd":  sf / sk  if sk else None,
            "bovina_usd": bf / bk  if bk else None,
            "frango_usd": ff / fk  if fk else None,
        }

    log("Preenchendo planilha...")
    meses_ok = []

    for key in meses_para_buscar:
        if key not in date_rows:
            log(f"  {key[1]:02d}/{key[0]}: linha nao encontrada — pulando")
            continue

        rn = date_rows[key]
        d  = dados_final[key]

        if all(v is None for v in (d["suina_kt"], d["bovina_kt"], d["frango_kt"])):
            log(f"  {key[1]:02d}/{key[0]}: sem dados na API — pulando")
            continue

        if last_ref_row is None or last_ref_row >= rn:
            col_d = sh.range(f"D2:D{rn - 1}").value
            if not isinstance(col_d, list): col_d = [col_d]
            for i in range(len(col_d) - 1, -1, -1):
                if col_d[i] is not None:
                    last_ref_row = i + 2
                    break

        if last_ref_row:
            sh.range(f"D{last_ref_row}:O{last_ref_row}").api.Copy()
            sh.range(f"D{rn}:O{rn}").api.PasteSpecial(Paste=-4122)
            sh.book.app.api.CutCopyMode = False

        sh.range(f"D{rn}").value = d["suina_kt"]
        sh.range(f"E{rn}").value = d["bovina_kt"]
        sh.range(f"F{rn}").value = d["frango_kt"]
        sh.range(f"H{rn}").value = d["suina_usd"]
        sh.range(f"J{rn}").value = d["bovina_usd"]
        sh.range(f"L{rn}").value = d["frango_usd"]

        sh.range(f"I{rn}").formula = f"=H{rn}*$O{rn}"
        sh.range(f"K{rn}").formula = f"=J{rn}*$O{rn}"
        sh.range(f"M{rn}").formula = f"=L{rn}*$O{rn}"

        ref_formula = sh.range(f"O{last_ref_row}").formula
        if ref_formula and ref_formula.startswith("="):
            sh.range(f"O{rn}").formula = adjust_formula_row(ref_formula, last_ref_row, rn)
        else:
            sh.range(f"O{rn}").value = sh.range(f"O{last_ref_row}").value

        log(f"  OK {key[1]:02d}/{key[0]}")
        meses_ok.append(key)
        last_ref_row = rn

    return meses_ok, dados_final


if __name__ == "__main__":
    main()
