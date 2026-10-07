"""Pipeline cloud Transportes com resumo, um download e um upload."""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from cloud_pipeline import Step, run

STEPS = [
    # Comex Stat ja e usado pelo BeefBR e pelo WEG a partir do runner, entao
    # este e o passo confiavel do pipeline.
    Step("SECEX Graos", "Transportes/update_secex.py"),
    # O IMEA exige login (IMEA_USERNAME / IMEA_PASSWORD) e o api1.imea.com.br
    # nunca foi testado a partir de um runner do GitHub -- portal privado pode
    # recusar IP de nuvem. Se cair, o SECEX ja coletado ainda sobe e o frete
    # continua vindo da execucao local ate se confirmar que a fonte responde.
    # Ver o workflow "Diag conectividade". Desde 07/10/2026 a falha deixa o
    # job vermelho, em vez de so um aviso no resumo.
    Step("Frete de graos (IMEA)", "Transportes/imea_frete_extractor.py"),
]

if __name__ == "__main__":
    run("Transportes", "Transportes.xlsm", STEPS)
