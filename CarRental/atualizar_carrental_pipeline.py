"""Pipeline cloud CarRental com resumo, um download e um upload."""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from cloud_pipeline import Step, run

STEPS = [
    # IPCA de automovel novo e usado, emendando 4 tabelas do SIDRA (IBGE). O
    # apisidra.ibge.gov.br ja e usado pelo BeefBR e pelo FrangoBR a partir do
    # runner, entao a fonte e confiavel aqui.
    Step("IPCA carros (SIDRA)", "CarRental/atualiza_sidra_carrental.py"),
]

if __name__ == "__main__":
    run("CarRental", "CarRental.xlsm", STEPS)
