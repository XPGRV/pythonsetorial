"""Pipeline cloud WEG com resumo, um download e um upload."""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from cloud_pipeline import Step, run

STEPS = [
    Step("SECEX Transformadores", "WEG/update_secex_transformadores.py"),
    Step("SECEX Motores", "WEG/update_secex_motores.py"),
    # O FRED as vezes bloqueia requisicao automatizada (WAF/Akamai) ou devolve
    # 504; se cair, o pipeline segue e o SECEX ja coletado ainda e' enviado.
    Step("PPI Transformadores (FRED)", "WEG/atualiza_transformadores.py", optional=True),
]
if __name__ == "__main__":
    run("WEG", "WEG - Setorial.xlsm", STEPS)
