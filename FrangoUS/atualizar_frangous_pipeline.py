"""Pipeline cloud FrangoUS com um download e um upload."""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from cloud_pipeline import Step, run

STEPS = [
    Step("Production", "FrangoUS/atualiza_production_frangous.py"),
    Step("Chicks hatched", "FrangoUS/atualiza_production_frangous_chicks_hatched.py"),
    Step("Abates USDA", "FrangoUS/atualiza_abates_frangous_usda.py"),
    Step("Rate of lay", "FrangoUS/atualiza_production_frangous_rateolay.py"),
    Step("Intended placements", "FrangoUS/atualiza_intended_placements_frangous.py"),
    Step("USDA FrangoUS", "FrangoUS/atualiza_usda_frangous.py"),
    Step("Forecast ERS LDP", "FrangoUS/coletar_usda_ers_ldp_forecasts.py"),
    Step("Wholesale composite", "FrangoUS/atualiza_wholesale_broiler_composite.py"),
    Step("National composite", "FrangoUS/atualiza_national_composite.py"),
    Step("National composite price", "FrangoUS/atualiza_national_composite_price.py"),
    Step("Consolidador FrangoUS", "FrangoUS/atualizar_frangoUS.py"),
]
if __name__ == "__main__":
    run("FrangoUS", "FrangoUS.xlsm", STEPS)
