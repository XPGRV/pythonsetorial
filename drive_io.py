"""
drive_io.py — I/O de planilhas .xlsm via Google Drive (service account).

Funciona em dois ambientes, sem mudar o codigo do coletor:
  - LOCAL: le as credenciais de `service-account.json` (na raiz, fora do git).
  - NUVEM: le de variaveis de ambiente (GitHub Secrets):
        GOOGLE_SERVICE_ACCOUNT_EMAIL
        GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY   (com \\n literais, igual na Vercel)

Uso tipico num coletor:

    import drive_io, openpyxl
    fid = drive_io.find_file_id("FrangoUS.xlsm")
    caminho = drive_io.download_to_temp(fid)          # baixa p/ arquivo temporario
    wb = openpyxl.load_workbook(caminho, keep_vba=True, data_only=False)
    ...                                               # edita a planilha
    wb.save(caminho)
    drive_io.update_file(fid, caminho)                # sobe de volta (in-place)

Observacoes:
  - A service account ATUALIZA arquivos existentes, mas NAO cria (sem cota).
  - A pasta que contem as planilhas precisa estar compartilhada com o e-mail
    da service account como *Editor*.
"""

from __future__ import annotations

import io
import os
import tempfile
from pathlib import Path

import net  # noqa: F401 — ao importar, deixa todo `requests` resiliente (retry)

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

SCOPES = ["https://www.googleapis.com/auth/drive"]
MIME_XLSM = "application/vnd.ms-excel.sheet.macroEnabled.12"

_SA_FILE = Path(__file__).resolve().parent / "service-account.json"

# Cache do cliente para nao reautenticar a cada chamada.
_drive_client = None


def _credentials():
    """Credenciais: arquivo local se existir, senao variaveis de ambiente."""
    if _SA_FILE.exists():
        return service_account.Credentials.from_service_account_file(
            str(_SA_FILE), scopes=SCOPES
        )

    email = os.environ.get("GOOGLE_SERVICE_ACCOUNT_EMAIL")
    key = os.environ.get("GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY")
    if email and key:
        info = {
            "type": "service_account",
            "client_email": email,
            "private_key": key.replace("\\n", "\n"),
            "token_uri": "https://oauth2.googleapis.com/token",
        }
        return service_account.Credentials.from_service_account_info(info, scopes=SCOPES)

    raise RuntimeError(
        "Credenciais do Drive nao encontradas: coloque service-account.json na raiz "
        "ou defina GOOGLE_SERVICE_ACCOUNT_EMAIL e GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY."
    )


def drive():
    """Cliente da Drive API v3 (com cache)."""
    global _drive_client
    if _drive_client is None:
        _drive_client = build("drive", "v3", credentials=_credentials(), cache_discovery=False)
    return _drive_client


def find_file_id(name: str, folder_id: str | None = None) -> str:
    """ID do arquivo pelo nome (opcionalmente dentro de uma pasta). Erra se 0 ou >1."""
    escaped = name.replace("'", "\\'")
    q = f"name = '{escaped}' and trashed = false"
    if folder_id:
        q = f"'{folder_id}' in parents and " + q
    res = (
        drive()
        .files()
        .list(
            q=q,
            fields="files(id,name,parents,modifiedTime)",
            includeItemsFromAllDrives=True,
            supportsAllDrives=True,
            pageSize=10,
        )
        .execute()
    )
    files = res.get("files", [])
    if not files:
        raise FileNotFoundError(f"Planilha nao encontrada no Drive: {name}")
    if len(files) > 1:
        locais = ", ".join(f"{f['id']}" for f in files)
        raise RuntimeError(f"Mais de uma '{name}' encontrada (ids: {locais}). Passe folder_id.")
    return files[0]["id"]


def download_to_temp(file_id: str, suffix: str = ".xlsm") -> str:
    """Baixa o arquivo para um temporario e devolve o caminho."""
    buf = io.BytesIO()
    req = drive().files().get_media(fileId=file_id, supportsAllDrives=True)
    dl = MediaIoBaseDownload(buf, req)
    done = False
    while not done:
        _, done = dl.next_chunk()
    fd, path = tempfile.mkstemp(suffix=suffix)
    with os.fdopen(fd, "wb") as fh:
        fh.write(buf.getvalue())
    return path


def update_file(file_id: str, local_path: str, mime: str = MIME_XLSM) -> dict:
    """Sobrescreve o conteudo do arquivo existente no Drive (in-place)."""
    media = MediaFileUpload(local_path, mimetype=mime, resumable=False)
    return (
        drive()
        .files()
        .update(
            fileId=file_id,
            media_body=media,
            fields="id,name,size,modifiedTime,md5Checksum",
            supportsAllDrives=True,
        )
        .execute()
    )
