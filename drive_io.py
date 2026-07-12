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
import time
from pathlib import Path

import net  # noqa: F401 — ao importar, deixa todo `requests` resiliente (retry)

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

try:
    import httplib2
    import google_auth_httplib2
except ImportError:  # pragma: no cover - fallback para ambiente local antigo
    httplib2 = None
    google_auth_httplib2 = None

SCOPES = ["https://www.googleapis.com/auth/drive"]
MIME_XLSM = "application/vnd.ms-excel.sheet.macroEnabled.12"
DRIVE_TIMEOUT_S = int(os.environ.get("DRIVE_TIMEOUT_S", "120"))

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
        creds = _credentials()
        if httplib2 is not None and google_auth_httplib2 is not None:
            http = google_auth_httplib2.AuthorizedHttp(
                creds,
                http=httplib2.Http(timeout=DRIVE_TIMEOUT_S),
            )
            _drive_client = build(
                "drive",
                "v3",
                http=http,
                cache_discovery=False,
            )
        else:
            _drive_client = build(
                "drive",
                "v3",
                credentials=creds,
                cache_discovery=False,
            )
    return _drive_client


def find_file_id(name: str, folder_id: str | None = None) -> str:
    """ID do arquivo pelo nome (opcionalmente dentro de uma pasta). Erra se 0 ou >1."""
    print(f"[Drive] Localizando {name!r}...", flush=True)
    t0 = time.monotonic()
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
    print(f"[Drive] {name!r} encontrada em {time.monotonic() - t0:.1f}s.", flush=True)
    return files[0]["id"]


def download_to_temp(file_id: str, suffix: str = ".xlsm") -> str:
    """Baixa o arquivo para um temporario e devolve o caminho."""
    print("[Drive] Baixando planilha para arquivo temporario...", flush=True)
    t0 = time.monotonic()
    buf = io.BytesIO()
    req = drive().files().get_media(fileId=file_id, supportsAllDrives=True)
    dl = MediaIoBaseDownload(buf, req)
    done = False
    while not done:
        status, done = dl.next_chunk()
        if status:
            print(f"[Drive] Download {int(status.progress() * 100)}%.", flush=True)
    fd, path = tempfile.mkstemp(suffix=suffix)
    with os.fdopen(fd, "wb") as fh:
        fh.write(buf.getvalue())
    tamanho_mb = len(buf.getvalue()) / (1024 * 1024)
    print(
        f"[Drive] Download concluido ({tamanho_mb:.1f} MB) em {time.monotonic() - t0:.1f}s.",
        flush=True,
    )
    return path


def _fix_comment_paths(path: str) -> None:
    """O openpyxl salva comentarios em xl/comments/commentN.xml — caminho que o
    Excel trata como 'conteudo com problema' (pede recuperacao ao abrir, e pode
    travar o macro do Bloomberg). Move de volta pro padrao xl/commentsN.xml e
    corrige as referencias (.rels e [Content_Types]). No-op se nao houver."""
    import re
    import shutil
    import tempfile
    import zipfile

    with zipfile.ZipFile(path, "r") as zin:
        names = zin.namelist()
        if not any(re.match(r"xl/comments/comment\d+\.xml$", n) for n in names):
            return  # nada a corrigir
        fd, tmp = tempfile.mkstemp(suffix=".xlsm")
        os.close(fd)
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                m = re.match(r"xl/comments/comment(\d+)\.xml$", item.filename)
                if m:
                    zout.writestr(f"xl/comments{m.group(1)}.xml", data)
                    continue
                if item.filename.endswith(".rels") or item.filename == "[Content_Types].xml":
                    txt = data.decode("utf-8")
                    txt = re.sub(r"/xl/comments/comment(\d+)\.xml", r"/xl/comments\1.xml", txt)
                    txt = re.sub(r"(\.\./)?comments/comment(\d+)\.xml", r"\1comments\2.xml", txt)
                    data = txt.encode("utf-8")
                zout.writestr(item, data)
    shutil.move(tmp, path)


def _fix_single_cell_array_formulas(path: str) -> None:
    """Remove metadados de formula matricial de 1 celula.

    O openpyxl pode preservar/gravar formulas como:
        <f t="array" ref="T5">...</f>
    Quando esse padrao aparece em muitas celulas individuais, o Excel pode abrir
    a planilha com reparo:
        "Repaired Records: Cell information from /xl/worksheets/sheetN.xml"

    Para formulas cujo `ref` e uma unica celula, removemos apenas os atributos
    `t="array"` e `ref="..."`. A formula em si permanece.
    """
    import re
    import shutil
    import tempfile
    import zipfile

    cell_re = re.compile(
        r'(<c\b[^>]*\br="([^"]+)"[^>]*>.*?</c>)',
        flags=re.DOTALL,
    )
    formula_re = re.compile(r"<f\b([^>]*)>", flags=re.DOTALL)

    def normalize_cell(match: re.Match) -> str:
        cell_xml = match.group(1)
        cell_ref = match.group(2)
        f_match = formula_re.search(cell_xml)
        if not f_match:
            return cell_xml

        attrs = f_match.group(1)
        if not re.search(r'\bt="array"', attrs):
            return cell_xml

        ref_match = re.search(r'\bref="([^"]+)"', attrs)
        if not ref_match:
            return cell_xml

        formula_ref = ref_match.group(1)
        ref_unica_celula = re.match(r"^\$?[A-Za-z]{1,3}\$?\d+$", formula_ref)
        if formula_ref != cell_ref and not ref_unica_celula:
            return cell_xml

        attrs = re.sub(r'\s+t="array"', "", attrs)
        attrs = re.sub(r'\s+ref="[^"]+"', "", attrs)
        novo_f = f"<f{attrs}>"
        return cell_xml[: f_match.start()] + novo_f + cell_xml[f_match.end() :]

    mudou = False
    with zipfile.ZipFile(path, "r") as zin:
        fd, tmp = tempfile.mkstemp(suffix=".xlsm")
        os.close(fd)
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if re.match(r"xl/worksheets/sheet\d+\.xml$", item.filename):
                    txt = data.decode("utf-8")
                    novo = cell_re.sub(normalize_cell, txt)
                    if novo != txt:
                        mudou = True
                        data = novo.encode("utf-8")
                zout.writestr(item, data)

    if mudou:
        shutil.move(tmp, path)
    else:
        os.remove(tmp)


def update_file(file_id: str, local_path: str, mime: str = MIME_XLSM) -> dict:
    """Sobrescreve o conteudo do arquivo existente no Drive (in-place)."""
    print("[Drive] Preparando upload da planilha...", flush=True)
    t0 = time.monotonic()
    _fix_comment_paths(local_path)
    _fix_single_cell_array_formulas(local_path)
    media = MediaFileUpload(local_path, mimetype=mime, resumable=False)
    res = (
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
    print(f"[Drive] Upload concluido em {time.monotonic() - t0:.1f}s.", flush=True)
    return res
