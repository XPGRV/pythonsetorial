"""Orquestrador: um download/upload e resumo por planilha/workflow.

UMA ETAPA QUEBRADA NAO DERRUBA AS OUTRAS
----------------------------------------
Ate 07/10/2026 a primeira etapa com erro abortava o pipeline, e como o
upload so acontece no fim, o trabalho das etapas que JA tinham rodado era
jogado fora. Quando o USDA renomeou uma aba, BeefUS e FrangoUS pararam de
atualizar por inteiro -- 3 e 11 etapas -- por causa de uma so. O mesmo valeu
para o SIDRA em BeefBR e FrangoBR.

Agora toda etapa roda, o upload acontece com o que deu certo, e o pipeline
falha no fim listando quem quebrou. O job continua vermelho: o que muda e'
que o estrago fica restrito a fonte com problema.

A flag `optional` deixou de decidir se o pipeline continua -- ela decide a
SEVERIDADE: etapa opcional e aviso, etapa normal e erro e deixa o job
vermelho.

Cuidado conhecido: as etapas nao sao independentes entre si (no FrangoBR o
copiar_secex alimenta as seguintes; os consolidadores rodam por ultimo). Se
uma etapa de base falhar, as seguintes trabalham com o dado da rodada
anterior em vez de nao rodar. E' uma troca deliberada: dado parcialmente
velho e melhor do que planilha inteira sem atualizar.
"""
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
    # Toda etapa que falha deixa o pipeline seguir (ver docstring). O que
    # `optional` muda e a severidade: True = aviso e job verde; False = erro
    # e job vermelho.
    #
    # Em 07/10/2026 nenhuma etapa usa optional=True. As tres que usavam (GATS,
    # FRED, IMEA) voltaram a ser normais a pedido: com o aviso ficando so no
    # resumo do run, nao dava para distinguir fonte intermitente de fonte
    # morta -- foi assim que o SIDRA passou duas semanas quebrado sem ninguem
    # ver. O campo continua funcionando para quem precisar dele de novo.
    optional: bool = False

def _summary(text: str) -> None:
    path = os.environ.get("GITHUB_STEP_SUMMARY")
    if path:
        with open(path, "a", encoding="utf-8") as file:
            file.write(text + "\n")

def _table_text(value: str) -> str:
    return value.replace("|", "\\|").replace("\n", " ").strip()


def _annotation_text(value: str) -> str:
    """Uma linha so: quebra de linha encerraria a anotacao no meio."""
    return " ".join(value.split())

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
    falhas: list[str] = []
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
                    f"| {_table_text(step.name)} | "
                    f"{'⚠️ Falhou' if step.optional else '❌ Erro'} | "
                    f"{_table_text(ultima_linha)} |"
                )
                if not step.optional:
                    falhas.append(f"{step.name} ({step.script}, exit {returncode})")
                    aviso_erro = _annotation_text(
                        f"{name}: a etapa '{step.name}' falhou (exit code "
                        f"{returncode}). O pipeline seguiu para as demais e o "
                        f"upload acontece com o que deu certo, mas essa fonte "
                        f"NAO atualizou nesta rodada."
                    )
                    if os.environ.get("GITHUB_ACTIONS"):
                        print(f"::error title=Coleta {name}::{aviso_erro}", flush=True)
                    print(f"[erro] {aviso_erro}", flush=True)
                    if index < len(steps):
                        time.sleep(2)
                    continue
                # Passo opcional falhando deixa o job VERDE, e o aviso ficaria
                # so na tabela do Resumo -- facil de nunca ver. A validacao do
                # dashboard tambem nao pega: ela conta pontos das series, e as
                # linhas historicas continuam la, entao os minimos seguem
                # satisfeitos mesmo com a fonte parada ha meses. O ::warning::
                # sobe o aviso para o topo do run no Actions.
                aviso = _annotation_text(
                    f"{name}: a etapa opcional '{step.name}' falhou "
                    f"(exit code {returncode}). O pipeline seguiu e o upload "
                    f"aconteceu, mas essa fonte NAO atualizou nesta rodada."
                )
                if os.environ.get("GITHUB_ACTIONS"):
                    print(f"::warning title=Coleta {name}::{aviso}", flush=True)
                print(f"[aviso] {aviso}", flush=True)
            else:
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

        # O upload ja aconteceu: o que deu certo esta no Drive. Falhar aqui
        # serve para o job ficar vermelho e alguem olhar.
        if falhas:
            raise RuntimeError(
                f"{len(falhas)} de {len(steps)} etapas falharam (o upload do "
                f"que deu certo foi feito): " + "; ".join(falhas)
            )
    finally:
        try:
            os.remove(path)
        except FileNotFoundError:
            pass
