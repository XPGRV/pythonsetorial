import re
import sys
from copy import copy
from datetime import datetime
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import drive_io

# Caminhos dos arquivos
DRIVE_ROOT = Path(
    r"G:\.shortcut-targets-by-id\11q1ngdqbySnDMCQSP4fiwPztT6ju2F4Z"
)
BEEF_NAME   = "BeefBR.xlsm"
FRANGO_NAME = "FrangoBR.xlsm"
SHEET_NAME  = "SECEX"


def localizar_planilha(nome: str) -> Path:
    """Localiza uma planilha sem depender dos nomes das pastas intermediárias."""
    encontrados = list(DRIVE_ROOT.rglob(nome))

    if not encontrados:
        raise FileNotFoundError(
            f"Planilha {nome!r} não encontrada dentro de {DRIVE_ROOT}"
        )

    if len(encontrados) > 1:
        caminhos = "\n".join(f"  - {path}" for path in encontrados)
        raise RuntimeError(
            f"Foram encontradas várias planilhas chamadas {nome!r}:\n{caminhos}"
        )

    return encontrados[0]


def log(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")

def corrigir_formula(formula):
    """
    Remove vínculos externos e ajusta colunas da BBG_Dados (G -> I).
    """
    if not isinstance(formula, str) or not formula.startswith("="):
        return formula
    
    # Remove referências ao arquivo original [BeefBR.xlsm]
    formula = formula.replace("[BeefBR.xlsm]", "").replace("'[BeefBR.xlsm]'", "")
    
    # Ajusta colunas da aba BBG_Dados se necessário
    if "BBG_Dados" in formula:
        # Troca G por I logo após o nome da aba
        formula = re.sub(r"(['\"]?BBG_Dados['\"]?!\$?)G(?![a-zA-Z])", r"\1I", formula, flags=re.IGNORECASE)
        # Troca G por I após o sinal de range (:)
        formula = re.sub(r"(?<=[:])\$?G(?![a-zA-Z])", r"I", formula, flags=re.IGNORECASE)
        
    return formula


def copy_sheet(src_ws, dst_ws):
    """Copia a aba inteira (valores/formulas + formato) da fonte p/ o destino,
    corrigindo as formulas. Equivale ao Cells.Clear + UsedRange.Copy do Excel."""
    # Desfaz merges do destino (senao nao da pra limpar/escrever nas celulas).
    for rng in list(dst_ws.merged_cells.ranges):
        dst_ws.unmerge_cells(str(rng))
    # Limpa valores do destino.
    for row in dst_ws.iter_rows():
        for cell in row:
            cell.value = None
    # Copia cada celula da fonte (valor corrigido + formato).
    for row in src_ws.iter_rows():
        for cell in row:
            dst = dst_ws.cell(row=cell.row, column=cell.column)
            dst.value = corrigir_formula(cell.value)
            dst.font = copy(cell.font)
            dst.fill = copy(cell.fill)
            dst.border = copy(cell.border)
            dst.alignment = copy(cell.alignment)
            dst.number_format = cell.number_format
    # Larguras de coluna e merges.
    for col, dim in src_ws.column_dimensions.items():
        if dim.width:
            dst_ws.column_dimensions[col].width = dim.width
    for rng in list(src_ws.merged_cells.ranges):
        dst_ws.merge_cells(str(rng))


def main():
    log("Iniciando copia da aba SECEX (BeefBR -> FrangoBR) via Drive...")
    beef_fid = drive_io.find_file_id(BEEF_NAME)
    frango_fid = drive_io.find_file_id(FRANGO_NAME)
    beef_path = drive_io.download_to_temp(beef_fid)
    frango_path = drive_io.download_to_temp(frango_fid)

    wb_beef = openpyxl.load_workbook(beef_path, data_only=False)
    wb_frango = openpyxl.load_workbook(frango_path, keep_vba=True, data_only=False)
    try:
        log(f"Copiando aba '{SHEET_NAME}' e corrigindo formulas...")
        copy_sheet(wb_beef[SHEET_NAME], wb_frango[SHEET_NAME])
        wb_frango.save(frango_path)
        drive_io.update_file(frango_fid, frango_path)
        log("FrangoBR salvo no Drive. Concluido.")
    finally:
        wb_frango.close()
        wb_beef.close()


if __name__ == "__main__":
    main()
