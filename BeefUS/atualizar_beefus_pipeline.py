"""Pipeline cloud BeefUS com um download e um upload."""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from cloud_pipeline import Step, run

STEPS = [
    Step("Abates USDA", "BeefUS/atualiza_abates_usda.py"),
    Step("Preco bezerro USDA NASS", "BeefUS/atualiza_px_bezerro_usa.py"),
    Step("Forecasts ERS LDP", "BeefUS/coletar_usda_ers_ldp_forecasts.py"),
    Step("Consolidador BeefUS", "BeefUS/atualizar_beefus.py"),
]
if __name__ == "__main__":
    run("BeefUS", "BeefUS.xlsm", STEPS)

