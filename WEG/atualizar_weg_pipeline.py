"""Pipeline cloud WEG com resumo, um download e um upload."""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from cloud_pipeline import Step, run

STEPS = [
    Step("SECEX Transformadores", "WEG/update_secex_transformadores.py"),
    Step("SECEX Motores", "WEG/update_secex_motores.py"),
    # O FRED as vezes bloqueia requisicao automatizada (WAF/Akamai) ou devolve
    # 504. Se cair, o pipeline segue e o SECEX ja coletado ainda e' enviado --
    # isso vale para qualquer etapa desde 07/10/2026. A diferenca e que agora
    # a falha deixa o job vermelho, para dar de distinguir queda eventual de
    # fonte morta.
    Step("PPI Transformadores (FRED)", "WEG/atualiza_transformadores.py"),
]
if __name__ == "__main__":
    run("WEG", "WEG - Setorial.xlsm", STEPS)
