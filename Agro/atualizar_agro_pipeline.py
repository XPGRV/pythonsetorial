"""Pipeline cloud Agro com resumo, um download e um upload.

O Agro nao tem coletor proprio: os precos chegam pelo Bloomberg (BDH). O que
este pipeline faz e' deixar a planilha PRONTA para o terminal preencher --
escrever os contratos do ciclo na aba Futuros e arrastar as formulas de
desconto da BBG_Dados, que ficavam para tras e travavam os graficos.

Como os outros pipelines, roda ANTES do Bloomberg: o openpyxl grava as formulas
sem valor em cache e o RefreshAll+F9 de la repoe.
"""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from cloud_pipeline import Step, run

STEPS = [
    # Deixa a planilha pronta para o Bloomberg: os contratos do ciclo com as
    # formulas BDP/BDH ao lado, para o terminal so preencher.
    Step("Tickers dos futuros", "Agro/atualizar_futuros.py"),
    Step("Descontos (BBG_Dados)", "Agro/atualizar_agro.py"),
]

if __name__ == "__main__":
    run("Agro", "Agro.xlsm", STEPS)
