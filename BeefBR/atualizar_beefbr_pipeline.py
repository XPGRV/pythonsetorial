"""Executa todas as coletas BeefBR em sequencia.

Com ``--cloud``, baixa BeefBR.xlsm uma unica vez, compartilha o arquivo
temporario entre as etapas e faz um unico upload ao final.
"""

from __future__ import annotations

import os
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
    cloud_args: tuple[str, ...] = ()


STEPS: list[Step] = [
    Step("Abates (SIF)", "atualiza_abates.py"),
    Step("Abates (SIDRA)", "atualiza_abates_sidra.py"),
    Step("SECEX (Comex Stat)", "update_secex.py"),
    Step("SECEX (USDA GATS)", "atualizar_secex.py"),
    Step(
        "Preco Bovina (Balanca)",
        "atualiza_px_bovina_balanca.py",
        ("--cloud-beefbr",),
    ),
    Step("Consolidar (BeefBR)", "atualizar_beefbr.py"),
]

SLEEP_BETWEEN_STEPS_S = 2.0


def _print_banner(msg: str) -> None:
    bar = "=" * 72
    print(f"\n{bar}\n{msg}\n{bar}", flush=True)


def _run_step(
    step: Step,
    cwd: Path,
    *,
    cloud: bool,
    env: dict[str, str] | None,
) -> None:
    script_path = cwd / step.script
    if not script_path.exists():
        raise FileNotFoundError(f"Script nao encontrado: {script_path}")

    _print_banner(f"Iniciando: {step.name} ({step.script})")
    args = step.cloud_args if cloud else ()
    cmd = [sys.executable, str(script_path), *args]

    with subprocess.Popen(
        cmd,
        cwd=str(cwd),
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
        env=env,
    ) as proc:
        assert proc.stdout is not None
        for line in proc.stdout:
            print(line.rstrip("\n"), flush=True)

        rc = proc.wait()
        if rc != 0:
            raise RuntimeError(f"Etapa falhou (exit code {rc}): {step.script}")

    print(f"\nConcluido: {step.name}", flush=True)


def run_all(steps: Iterable[Step] = STEPS, *, cloud: bool = False) -> None:
    steps = list(steps)
    cwd = Path(__file__).resolve().parent
    root = cwd.parent

    _print_banner("Pipeline BeefBR - execucao sequencial")
    print(f"Pasta: {cwd}", flush=True)
    print(f"Python: {sys.executable}", flush=True)
    print(f"Etapas: {len(steps)}", flush=True)

    shared_path: str | None = None
    file_id: str | None = None
    child_env: dict[str, str] | None = None

    if cloud:
        sys.path.insert(0, str(root))
        import drive_io

        file_id = drive_io.find_file_id("BeefBR.xlsm")
        shared_path = drive_io.download_to_temp(file_id)
        child_env = os.environ.copy()
        child_env[drive_io.SHARED_PATH_ENV] = shared_path
        child_env[drive_io.SHARED_FILE_ID_ENV] = file_id
        child_env[drive_io.SHARED_NAME_ENV] = "BeefBR.xlsm"

    try:
        for index, step in enumerate(steps, start=1):
            print(f"\n[{index}/{len(steps)}] {step.name}", flush=True)
            _run_step(step, cwd, cloud=cloud, env=child_env)
            if index < len(steps) and SLEEP_BETWEEN_STEPS_S > 0:
                time.sleep(SLEEP_BETWEEN_STEPS_S)

        if cloud and file_id and shared_path:
            print("\n[Drive] Enviando resultado final do pipeline...", flush=True)
            drive_io.update_file(file_id, shared_path)
    finally:
        if shared_path:
            try:
                os.remove(shared_path)
            except FileNotFoundError:
                pass

    _print_banner("Pipeline finalizado com sucesso")


def main() -> None:
    try:
        run_all(cloud="--cloud" in sys.argv)
    except KeyboardInterrupt:
        print("\nInterrompido pelo usuario.")
        raise SystemExit(130)


if __name__ == "__main__":
    main()
