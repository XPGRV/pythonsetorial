import re
import xlwings as xw
from datetime import datetime
from pathlib import Path

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

def main():
    log("Iniciando atualização da aba SECEX...")
    beef_path = localizar_planilha(BEEF_NAME)
    frango_path = localizar_planilha(FRANGO_NAME)
    app = xw.App(visible=False, add_book=False)
    try:
        app.api.AskToUpdateLinks = False
        app.display_alerts       = False
        
        # Abre as planilhas
        wb_beef   = app.books.open(str(beef_path), update_links=0, read_only=True)
        wb_frango = app.books.open(str(frango_path), update_links=0, read_only=False)

        sh_beef   = wb_beef.sheets[SHEET_NAME]
        sh_frango = wb_frango.sheets[SHEET_NAME]

        # 1. Atualizar conteúdo sem deletar a aba (Preserva links de outras abas)
        log(f"Limpando e sobrescrevendo dados na aba '{SHEET_NAME}'...")
        sh_frango.api.Cells.Clear() 
        sh_beef.api.UsedRange.Copy(Destination=sh_frango.api.Range("A1"))
        
        # 2. Corrigir fórmulas internas da nova SECEX
        log("Corrigindo fórmulas e removendo vínculos externos...")
        used = sh_frango.used_range
        formulas = used.formula
        
        if formulas:
            # Processa a matriz de fórmulas para ser mais rápido
            novas_formulas = []
            for row in formulas:
                nova_row = [corrigir_formula(cell) for cell in row]
                novas_formulas.append(nova_row)
            
            used.formula = novas_formulas

        log("Salvando alterações no FrangoBR...")
        wb_frango.save()
        wb_frango.close()
        wb_beef.close()
        log("Processo concluído com sucesso!")

    except Exception as e:
        log(f"ERRO: {e}")
    finally:
        app.quit()

if __name__ == "__main__":
    main()
