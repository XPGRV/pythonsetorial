"""Orquestrador: um download/upload e resumo por planilha/workflow."""
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

def _summary(text: str) -> None:
    path = os.environ.get("GITHUB_STEP_SUMMARY")
    if path:
        with open(path, "a", encoding="utf-8") as file:
            file.write(text + "\n")

def _table_text(value: str) -> str:
    return value.replace("|", "\\|").replace("\n", " ").strip()

def run(name: str, workbook: str, steps: list[Step]) -> None:
    root = Path(__file__).resolve().parent
    print(f"\n{'=' * 72}\nPipeline {name}: {len(steps)} etapas\n{'=' * 72}", flush=True)
    _summary(f"## Coleta {name}")
    _summary("")
    _summary("| Etapa | Status | Resultado |")
    _summary("|---|---:|---|")
    try:
        file_id = drive_io.find_file_id(workbook)
        path = drive_io.download_to_temp(file_id)
    except Exception as exc:
        _summary(f"| Download inicial | ❌ Erro | {_table_text(str(exc))} |")
        raise
    _summary(f"| Download inicial | ✅ OK | {_table_text(workbook)} |")
    env = os.environ.copy()
    env[drive_io.SHARED_PATH_ENV] = path
    env[drive_io.SHARED_FILE_ID_ENV] = file_id
    env[drive_io.SHARED_NAME_ENV] = workbook
    try:
        for index, step in enumerate(steps, 1):
            script = root / step.script
            print(f"\n[{index}/{len(steps)}] {step.name}", flush=True)
            ultima_linha = "Sem mensagem final"
            with subprocess.Popen(
                [sys.executable, "-u", str(script), *step.args],
                cwd=str(script.parent),
                env=env,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
            ) as proc:
                assert proc.stdout is not None
                for line in proc.stdout:
                    print(line, end="", flush=True)
                    if line.strip():
                        ultima_linha = line.strip()
                returncode = proc.wait()

            if returncode:
                _summary(
                    f"| {_table_text(step.name)} | ❌ Erro | "
                    f"{_table_text(ultima_linha)} |"
                )
                raise RuntimeError(
                    f"Etapa falhou (exit code {returncode}): {step.script}"
                )
            _summary(
                f"| {_table_text(step.name)} | ✅ OK | "
                f"{_table_text(ultima_linha)} |"
            )
            if index < len(steps):
                time.sleep(2)
        print("\n[Drive] Enviando resultado final do pipeline...", flush=True)
        try:
            upload = drive_io.update_file(file_id, path)
        except Exception as exc:
            _summary(f"| Upload final | ❌ Erro | {_table_text(str(exc))} |")
            raise
        _summary(
            "| Upload final | ✅ OK | "
            f"{_table_text(upload.get('name', workbook))} enviado ao Drive |"
        )
    finally:
        try:
            os.remove(path)
        except FileNotFoundError:
            pass
