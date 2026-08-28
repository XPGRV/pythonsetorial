"""Pipeline cloud Agro com resumo, um download e um upload.

O Agro nao tem coletor proprio: os precos chegam pelo Bloomberg (BDH) na aba
BBG_Dados. O que falta e' arrastar as formulas de desconto, que ficam para tras
e travam os graficos de desconto da dashboard.

Como os outros pipelines, roda ANTES do Bloomberg: o openpyxl grava as formulas
sem valor em cache e o RefreshAll+F9 de la repoe.
"""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from cloud_pipeline import Step, run

STEPS = [
    Step("Descontos (BBG_Dados)", "Agro/atualizar_agro.py"),
]

if __name__ == "__main__":
    run("Agro", "Agro.xlsm", STEPS)
