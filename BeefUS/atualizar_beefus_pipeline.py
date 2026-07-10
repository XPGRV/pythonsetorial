"""
atualizar_beefus_pipeline.py

Executa a atualização completa do BeefUS na ordem:
  1) atualiza_abates_usda.py
  2) atualiza_px_bezerro_usa.py
  3) coletar_usda_ers_ldp_forecasts.py
  4) atualizar_beefus.py

Roda cada etapa em um processo separado (subprocess), para evitar:
  - vazamento de estado entre scripts
  - conflitos de imports/globais
  - reaproveitamento acidental de instâncias do Excel

Uso:
  python atualizar_beefus_pipeline.py
"""

from __future__ import annotations

import subprocess
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable


@dataclass(frozen=True)
class Step:
    name: str
    script: str


STEPS: list[Step] = [
    Step("Abates (USDA ERS)", "atualiza_abates_usda.py"),
    Step("Preço bezerro (USDA NASS)", "atualiza_px_bezerro_usa.py"),
    Step("Forecasts ERS (LDP Outlook)", "coletar_usda_ers_ldp_forecasts.py"),
    Step("Consolidar (BeefUS)", "atualizar_beefus.py"),
]

# Pausa curta entre etapas para dar tempo do Excel liberar arquivo/COM.
SLEEP_BETWEEN_STEPS_S = 2.0


def _print_banner(msg: str) -> None:
    bar = "=" * 72
    print(f"\n{bar}\n{msg}\n{bar}", flush=True)


def _run_step(step: Step, cwd: Path) -> None:
    script_path = cwd / step.script
    if not script_path.exists():
        raise FileNotFoundError(f"Script não encontrado: {script_path}")

    _print_banner(f"Iniciando: {step.name}  ({step.script})")
    cmd = [sys.executable, "-u", str(script_path)]

    # Streama stdout+stderr em tempo real.
    with subprocess.Popen(
        cmd,
        cwd=str(cwd),
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
    ) as proc:
        assert proc.stdout is not None
        for line in proc.stdout:
            print(line.rstrip("\n"), flush=True)

        rc = proc.wait()
        if rc != 0:
            raise RuntimeError(f"Etapa falhou (exit code {rc}): {step.script}")

    print(f"\nConcluído: {step.name}")


def run_all(steps: Iterable[Step] = STEPS) -> None:
    cwd = Path(__file__).resolve().parent

    _print_banner("Pipeline BeefUS — execução sequencial")
    print(f"Pasta: {cwd}", flush=True)
    print(f"Python: {sys.executable}", flush=True)
    print(f"Etapas: {len(list(steps))}", flush=True)

    steps_list = list(steps)
    for i, step in enumerate(steps_list, start=1):
        print(f"\n[{i}/{len(steps_list)}] {step.name}", flush=True)
        _run_step(step, cwd=cwd)
        if i < len(steps_list) and SLEEP_BETWEEN_STEPS_S > 0:
            time.sleep(SLEEP_BETWEEN_STEPS_S)

    _print_banner("Pipeline finalizado com sucesso")


def main() -> None:
    try:
        run_all()
    except KeyboardInterrupt:
        print("\nInterrompido pelo usuário.")
        raise SystemExit(130)


if __name__ == "__main__":
    main()

