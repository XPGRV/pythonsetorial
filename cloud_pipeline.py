"""Orquestrador: um download e um upload por planilha/workflow."""
from __future__ import annotations
import os, subprocess, sys, time
from dataclasses import dataclass
from pathlib import Path

import drive_io

@dataclass(frozen=True)
class Step:
    name: str
    script: str
    args: tuple[str, ...] = ()

def run(name: str, workbook: str, steps: list[Step]) -> None:
    root = Path(__file__).resolve().parent
    print(f"\n{'=' * 72}\nPipeline {name}: {len(steps)} etapas\n{'=' * 72}", flush=True)
    file_id = drive_io.find_file_id(workbook)
    path = drive_io.download_to_temp(file_id)
    env = os.environ.copy()
    env[drive_io.SHARED_PATH_ENV] = path
    env[drive_io.SHARED_FILE_ID_ENV] = file_id
    env[drive_io.SHARED_NAME_ENV] = workbook
    try:
        for index, step in enumerate(steps, 1):
            script = root / step.script
            print(f"\n[{index}/{len(steps)}] {step.name}", flush=True)
            proc = subprocess.run(
                [sys.executable, "-u", str(script), *step.args],
                cwd=str(script.parent), env=env, check=False,
            )
            if proc.returncode:
                raise RuntimeError(
                    f"Etapa falhou (exit code {proc.returncode}): {step.script}"
                )
            if index < len(steps):
                time.sleep(2)
        print("\n[Drive] Enviando resultado final do pipeline...", flush=True)
        drive_io.update_file(file_id, path)
    finally:
        try:
            os.remove(path)
        except FileNotFoundError:
            pass
