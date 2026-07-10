"""
atualizar_selic.py

Preenche a coluna D (CDI mensal) da aba BBG_Dados em Planilha - Selic.xlsm
para todos os meses que ja tem data na coluna C mas ainda nao tem valor na coluna D.

Formula: media dos valores diarios de col B para aquele mes + 0.10

Requer: pip install pywin32
"""

import sys
import datetime
from collections import defaultdict
from pathlib import Path

DRIVE_ROOT = Path(
    r"G:\.shortcut-targets-by-id\11q1ngdqbySnDMCQSP4fiwPztT6ju2F4Z"
)
XLSX_NAME  = "Planilha - Selic.xlsm"
SHEET_NAME  = "BBG_Dados"
START_ROW   = 5
OFFSET_CDI  = 0.10   # CDI = media SELIC diaria + 0.10

XL_UP          = -4162
XL_CALC_MANUAL = -4135
XL_CALC_AUTO   = -4105


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


def ym(dt) -> tuple:
    """Retorna (year, month) de um objeto datetime do win32com."""
    return dt.year, dt.month


def main():
    try:
        import win32com.client
    except ImportError:
        print("ERRO: pywin32 nao instalado. Execute: pip install pywin32")
        sys.exit(1)

    try:
        path = localizar_planilha()
    except (FileNotFoundError, RuntimeError) as exc:
        print(f"ERRO: {exc}")
        sys.exit(1)

    today = datetime.date.today()
    print(f"SELIC -- {today:%d/%m/%Y}\n")

    excel = win32com.client.DispatchEx("Excel.Application")
    try: excel.Visible        = False
    except Exception: pass
    try: excel.DisplayAlerts  = False
    except Exception: pass
    try: excel.ScreenUpdating = False
    except Exception: pass

    wb = None
    calc_original = None
    try:
        wb = excel.Workbooks.Open(
            str(path),
            UpdateLinks=0,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
            Notify=False,
            AddToMru=False,
        )

        if wb.ReadOnly:
            wb.Close(SaveChanges=False)
            wb = None
            raise RuntimeError(
                f"A planilha {path.name} esta aberta ou bloqueada por outro processo."
            )

        calc_original = excel.Calculation
        excel.Calculation = XL_CALC_MANUAL
        ws  = wb.Sheets(SHEET_NAME)

        # ── 1. Le dados diários (col A = data, col B = taxa) ─────────────────
        last_daily = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row
        daily_by_month = defaultdict(list)
        for r in range(START_ROW, last_daily + 1):
            d = ws.Cells(r, 1).Value
            v = ws.Cells(r, 2).Value
            if d is not None and v is not None:
                daily_by_month[ym(d)].append(float(v))

        if not daily_by_month:
            raise RuntimeError("Nenhum dado diario encontrado em col A/B.")

        # ── 2. Le datas da col C e valores da col D ───────────────────────────
        last_c = ws.Cells(ws.Rows.Count, 3).End(XL_UP).Row

        filled   = 0
        skipped  = 0
        no_data  = 0

        for r in range(START_ROW, last_c + 1):
            c_val = ws.Cells(r, 3).Value   # data mensal em col C
            d_val = ws.Cells(r, 4).Value   # CDI mensal em col D (pode ser None)

            if c_val is None:
                continue

            # Ja tem valor — nao sobreescreve
            if d_val is not None:
                skipped += 1
                continue

            # Nao tem valor ainda — calcula
            key = ym(c_val)
            daily_vals = daily_by_month.get(key)
            if not daily_vals:
                no_data += 1
                continue

            cdi_mensal = sum(daily_vals) / len(daily_vals) + OFFSET_CDI
            ws.Cells(r, 4).Value = cdi_mensal
            print(f"  Linha {r:3d}  {key[1]:02d}/{key[0]}  CDI = {cdi_mensal:.4f}  ({len(daily_vals)} dia(s))")
            filled += 1

        # ── 3. Salva ──────────────────────────────────────────────────────────
        if filled:
            wb.Save()

        # Impede o Excel de exibir um aviso de alteracoes pendentes ao fechar.
        wb.Saved = True
        wb.Close(SaveChanges=False)
        wb = None

        if calc_original is not None:
            excel.Calculation = calc_original

        if filled:
            print(f"\n{filled} mes(es) preenchido(s). Salvo: {path.name}")
        else:
            print(f"Col D ja estava atualizada ({skipped} mes(es) com valor). Nada alterado.")
            if no_data:
                print(f"  {no_data} mes(es) em col C sem dados diarios correspondentes (futuro?).")

    except Exception:
        import traceback
        traceback.print_exc()
        sys.exit(1)
    finally:
        if wb is not None:
            try:
                # Em caso de erro, descarta apenas as alteracoes desta execucao.
                wb.Saved = True
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if calc_original is not None:
            try:
                excel.Calculation = calc_original
            except Exception:
                pass
        try:
            excel.ScreenUpdating = True
            excel.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
