"""Abre, atualiza, salva e fecha as planilhas setoriais em sequencia.

Cada arquivo so e liberado quando o Excel termina os calculos e nao ha mais
mensagens temporarias do Bloomberg nas celulas usadas pela planilha.

Requer: pip install pywin32
"""

from __future__ import annotations

import datetime as dt
import time
from pathlib import Path


DRIVE_ROOT = Path(
    r"G:\.shortcut-targets-by-id\11q1ngdqbySnDMCQSP4fiwPztT6ju2F4Z"
)

ARQUIVOS = [
    "FrangoUS.xlsm",
    "BeefBR.xlsm",
    "BeefUS.xlsm",
    "FrangoBR.xlsm",
    "Planilha - Selic.xlsm",
]

INTERVALO_VERIFICACAO = 3
VERIFICACOES_ESTAVEIS = 3
TEMPO_MINIMO_ATUALIZACAO = 15

XL_DONE = 0
XL_VALUES = -4163
XL_PART = 2
MSO_AUTOMATION_SECURITY_LOW = 1
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3

MARCADORES_PENDENTES = (
    "requesting data",
    "loading",
    "retrieving",
    "getting data",
    "aguardando dados",
    "carregando",
    "#n/a request",
)


def localizar_planilha(nome: str) -> Path:
    encontrados = list(DRIVE_ROOT.rglob(nome))

    if not encontrados:
        raise FileNotFoundError(
            f"Planilha {nome!r} nao encontrada dentro de {DRIVE_ROOT}"
        )

    if len(encontrados) > 1:
        caminhos = "\n".join(f"  - {path}" for path in encontrados)
        raise RuntimeError(
            f"Foram encontradas varias planilhas chamadas {nome!r}:\n{caminhos}"
        )

    return encontrados[0]


def _iterar_valores(valor):
    if isinstance(valor, (tuple, list)):
        for item in valor:
            yield from _iterar_valores(item)
    else:
        yield valor


def contar_pendencias_bloomberg(wb) -> int:
    """Conta mensagens temporarias do Bloomberg nas areas usadas."""
    pendencias = 0

    for ws in wb.Worksheets:
        try:
            area = ws.UsedRange
            valores = area.Value
        except Exception:
            continue

        # Find tambem enxerga textos de erro exibidos pelo Excel que podem
        # chegar ao Python como codigos numericos, nao como strings.
        for busca in ("*Requesting Data*", "*Loading*"):
            try:
                if area.Find(
                    What=busca,
                    LookIn=XL_VALUES,
                    LookAt=XL_PART,
                    MatchCase=False,
                ) is not None:
                    pendencias += 1
            except Exception:
                pass

        for valor in _iterar_valores(valores):
            if not isinstance(valor, str):
                continue
            texto = valor.strip().lower()
            if any(marcador in texto for marcador in MARCADORES_PENDENTES):
                pendencias += 1

    return pendencias


def aguardar_conclusao(excel, wb) -> None:
    """Espera conclusao real; nao usa um prazo fixo para cortar a tarefa."""
    estaveis = 0
    inicio = time.monotonic()
    ultimo_aviso = 0.0
    f9_reforco_enviado = False

    while estaveis < VERIFICACOES_ESTAVEIS:
        calculo_pronto = excel.CalculationState == XL_DONE
        pendencias = contar_pendencias_bloomberg(wb)

        decorrido = time.monotonic() - inicio
        tempo_minimo_cumprido = decorrido >= TEMPO_MINIMO_ATUALIZACAO

        # O Bloomberg frequentemente so aceita o F9 depois que as formulas
        # entram em "Requesting Data". Reenvia uma vez nesse momento exato.
        if pendencias > 0 and not f9_reforco_enviado:
            print("  Bloomberg em Requesting Data; reenviando F9...")
            enviar_f9(excel, wb)
            f9_reforco_enviado = True
            estaveis = 0
            time.sleep(INTERVALO_VERIFICACAO)
            continue

        if calculo_pronto and pendencias == 0 and tempo_minimo_cumprido:
            estaveis += 1
        else:
            estaveis = 0

        if decorrido - ultimo_aviso >= 15:
            print(
                f"  Aguardando: calculo={'OK' if calculo_pronto else 'em andamento'}, "
                f"pendencias Bloomberg={pendencias}, {decorrido:.0f}s"
            )
            ultimo_aviso = decorrido

        if estaveis < VERIFICACOES_ESTAVEIS:
            time.sleep(INTERVALO_VERIFICACAO)

    print(f"  Atualizacao concluida e estavel em {time.monotonic() - inicio:.0f}s.")


def salvar_e_confirmar(wb) -> None:
    wb.Save()

    while not wb.Saved:
        time.sleep(1)

    print("  Salvamento confirmado pelo Excel.")


def enviar_f9(excel, wb) -> None:
    """Coloca o Excel em primeiro plano e envia F9 sem bloquear o Python."""
    import pyautogui
    import win32con
    import win32gui

    wb.Activate()
    hwnd = excel.Hwnd
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    win32gui.SetForegroundWindow(hwnd)
    time.sleep(1)
    pyautogui.press("f9")


def processar_arquivo(excel, path: Path) -> None:
    wb = None
    try:
        # Habilita macros somente durante a abertura deste arquivo conhecido.
        excel.AutomationSecurity = MSO_AUTOMATION_SECURITY_LOW
        wb = excel.Workbooks.Open(
            str(path),
            UpdateLinks=3,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
            Notify=False,
            AddToMru=False,
        )
        excel.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE

        if wb.ReadOnly:
            raise RuntimeError(f"{path.name} foi aberto somente para leitura.")

        try:
            wb.RefreshAll()
        except Exception:
            pass

        # Envia o primeiro F9; se surgir Requesting Data, o monitor reenvia.
        print("  Enviando F9 para iniciar o calculo Bloomberg...")
        enviar_f9(excel, wb)

        aguardar_conclusao(excel, wb)
        salvar_e_confirmar(wb)

        wb.Saved = True
        wb.Close(SaveChanges=False)
        wb = None
        print(f"  {path.name} concluido.")

    finally:
        excel.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
        if wb is not None:
            try:
                wb.Saved = True
                wb.Close(SaveChanges=False)
            except Exception:
                pass


def main() -> None:
    import win32com.client

    caminhos = [(nome, localizar_planilha(nome)) for nome in ARQUIVOS]

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.ScreenUpdating = True

    erros: list[tuple[str, str]] = []
    try:
        for indice, (nome, path) in enumerate(caminhos, start=1):
            print(
                f"\n[{indice}/{len(caminhos)}] "
                f"[{dt.datetime.now():%H:%M:%S}] Abrindo {nome}..."
            )
            try:
                processar_arquivo(excel, path)
            except Exception as exc:
                erros.append((nome, str(exc)))
                print(f"  ERRO em {nome}: {exc}")

    finally:
        try:
            excel.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
            excel.Quit()
        except Exception:
            pass

    if erros:
        detalhes = "\n".join(f"  - {nome}: {erro}" for nome, erro in erros)
        raise RuntimeError(f"Falha em {len(erros)} arquivo(s):\n{detalhes}")

    print("\nTodos os arquivos foram atualizados, salvos e fechados.")


if __name__ == "__main__":
    main()
