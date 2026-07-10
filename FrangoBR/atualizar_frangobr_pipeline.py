"""
atualizar_frangobr_pipeline.py

Executa a atualização completa do FrangoBR na ordem:
  1) atualiza_abates_frango.py
  2) atualiza_sidra_frango.py
  3) copiar_secex.py
  4) ipca_api_extract.py
  5) atualizar_frangobr.py

Roda cada etapa em um processo separado (subprocess), para evitar:
  - vazamento de estado entre scripts
  - conflitos de imports/globais
  - reaproveitamento acidental de instâncias do Excel

Uso:
  python atualizar_frangobr_pipeline.py
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
    Step("Abates Frango (SIF)",       "atualiza_abates_frango.py"),
    Step("Abates Frango (SIDRA)",     "atualiza_sidra_frango.py"),
    Step("SECEX (Copiar)",            "copiar_secex.py"),
    Step("IPCA (API)",                "ipca_api_extract.py"),
    Step("Consolidar (FrangoBR)",     "atualizar_frangobr.py"),
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
    cmd = [sys.executable, str(script_path)]

    # Streama stdout+stderr em tempo real (essencial para scripts longos).
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

    _print_banner("Pipeline FrangoBR — execução sequencial")
    print(f"Pasta: {cwd}", flush=True)
    print(f"Python: {sys.executable}", flush=True)
    print(f"Etapas: {len(STEPS)}", flush=True)

    for i, step in enumerate(STEPS, start=1):
        print(f"\n[{i}/{len(STEPS)}] {step.name}", flush=True)
        _run_step(step, cwd=cwd)
        if i < len(STEPS) and SLEEP_BETWEEN_STEPS_S > 0:
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
