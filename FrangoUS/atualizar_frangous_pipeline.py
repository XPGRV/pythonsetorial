"""
atualizar_frangous_pipeline.py

Executa a atualizacao completa do FrangoUS na ordem:
  1)  atualiza_wholesale_broiler_composite.py
  2)  atualiza_national_composite.py
  3)  atualiza_national_composite_price.py
  4)  atualiza_production_frangous.py
  5)  atualiza_production_frangous_rateolay.py
  6)  atualiza_intended_placements_frangous.py
  7)  atualiza_production_frangous_chicks_hatched.py
  8)  atualiza_abates_frangous_usda.py
  9)  atualiza_usda_frangous.py
  10) atualizar_frangoUS.py

Uso:
  python atualizar_frangous_pipeline.py
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
    Step("Wholesale Broiler Composite",         "atualiza_wholesale_broiler_composite.py"),
    Step("National Composite",                  "atualiza_national_composite.py"),
    Step("National Composite Price",            "atualiza_national_composite_price.py"),
    Step("Production FrangoUS",                 "atualiza_production_frangous.py"),
    Step("Production FrangoUS (Rate of Lay)",   "atualiza_production_frangous_rateolay.py"),
    Step("Intended Placements FrangoUS",        "atualiza_intended_placements_frangous.py"),
    Step("Production FrangoUS (Chicks Hatched)","atualiza_production_frangous_chicks_hatched.py"),
    Step("Abates FrangoUS (USDA)",              "atualiza_abates_frangous_usda.py"),
    Step("USDA FrangoUS",                       "atualiza_usda_frangous.py"),
    Step("Consolidar (FrangoUS)",               "atualizar_frangoUS.py"),
]

SLEEP_BETWEEN_STEPS_S = 2.0


def _print_banner(msg: str) -> None:
    bar = "=" * 72
    print(f"\n{bar}\n{msg}\n{bar}", flush=True)


def _run_step(step: Step, cwd: Path) -> None:
    script_path = cwd / step.script
    if not script_path.exists():
        raise FileNotFoundError(f"Script nao encontrado: {script_path}")

    _print_banner(f"Iniciando: {step.name}  ({step.script})")
    cmd = [sys.executable, str(script_path)]

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

    print(f"\nConcluido: {step.name}")


def run_all(steps: Iterable[Step] = STEPS) -> None:
    cwd = Path(__file__).resolve().parent

    _print_banner("Pipeline FrangoUS — execucao sequencial")
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
        print("\nInterrompido pelo usuario.")
        raise SystemExit(130)


if __name__ == "__main__":
    main()
