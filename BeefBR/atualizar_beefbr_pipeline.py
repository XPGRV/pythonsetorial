"""Pipeline cloud BeefBR com resumo, um download e um upload."""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from cloud_pipeline import Step, run

STEPS = [
    Step("Abates SIF", "BeefBR/atualiza_abates.py"),
    Step("Abates SIDRA", "BeefBR/atualiza_abates_sidra.py"),
    Step("SECEX Comex Stat", "BeefBR/update_secex.py"),
    Step("SECEX USDA GATS", "BeefBR/atualizar_secex.py"),
    Step(
        "Preco de exportacao bovina",
        "BeefBR/atualiza_px_bovina_balanca.py",
        ("--cloud-beefbr",),
    ),
    Step("Consolidador BeefBR", "BeefBR/atualizar_beefbr.py"),
]
if __name__ == "__main__":
    run("BeefBR", "BeefBR.xlsm", STEPS)
