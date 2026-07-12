"""Pipeline cloud FrangoBR com um download e um upload do FrangoBR."""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from cloud_pipeline import Step, run

STEPS = [
    Step("Copiar SECEX", "FrangoBR/copiar_secex.py"),
    Step("Abates SIF", "FrangoBR/atualiza_abates_frango.py"),
    Step("Abates SIDRA", "FrangoBR/atualiza_sidra_frango.py"),
    Step("IPCA processados", "FrangoBR/ipca_api_extract.py"),
    Step(
        "Preco exportacao",
        "BeefBR/atualiza_px_bovina_balanca.py",
        ("--cloud-frangobr",),
    ),
    Step("Consolidador FrangoBR", "FrangoBR/atualizar_frangobr.py"),
]
if __name__ == "__main__":
    run("FrangoBR", "FrangoBR.xlsm", STEPS)
