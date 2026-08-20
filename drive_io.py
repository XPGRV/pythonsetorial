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
DRIVE_NUM_RETRIES = int(os.environ.get("DRIVE_NUM_RETRIES", "5"))
# Chunk menor que a planilha (~4 MB) de proposito: com um unico chunk o upload
# fica mudo do inicio ao fim e, se a conexao cair no meio, o arquivo INTEIRO e'
# reenviado. Em pedacos de 1 MB o progresso aparece no log e a retomada refaz
# apenas o pedaco que falhou.
DRIVE_UPLOAD_CHUNK_SIZE = 1024 * 1024
SHARED_PATH_ENV = "DRIVE_IO_SHARED_PATH"
SHARED_FILE_ID_ENV = "DRIVE_IO_SHARED_FILE_ID"
SHARED_NAME_ENV = "DRIVE_IO_SHARED_NAME"

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
            http_base = httplib2.Http(timeout=DRIVE_TIMEOUT_S)
            # Upload resumable: o Google responde 308 (Resume Incomplete) a cada
            # pedaco, com header `Range` e SEM `Location`. O httplib2 lista 308
            # como redirect e aborta com RedirectMissingLocation antes que o
            # googleapiclient — que sabe tratar 308 — veja a resposta. Tirar 308
            # de redirect_codes devolve o 308 para a camada de cima.
            # So' aparece em upload multi-chunk (arquivo maior que o chunk).
            http_base.redirect_codes = httplib2.REDIRECT_CODES - {308}
            http = google_auth_httplib2.AuthorizedHttp(creds, http=http_base)
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
    shared_path = os.environ.get(SHARED_PATH_ENV)
    shared_file_id = os.environ.get(SHARED_FILE_ID_ENV)
    shared_name = os.environ.get(SHARED_NAME_ENV)
    if shared_path and shared_file_id and shared_name == name:
        print(f"[Drive] Usando arquivo compartilhado local para {name!r}.", flush=True)
        return shared_file_id

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
    shared_path = os.environ.get(SHARED_PATH_ENV)
    shared_file_id = os.environ.get(SHARED_FILE_ID_ENV)
    if shared_path and shared_file_id == file_id:
        print("[Drive] Reutilizando planilha temporaria compartilhada.", flush=True)
        return shared_path

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


def _drop_calc_chain(path: str) -> None:
    """Remove xl/calcChain.xml do pacote.

    O calcChain e um INDICE da ordem em que o Excel calcula as formulas -- cache,
    nao dado. O openpyxl preserva o indice do arquivo original, mas as celulas de
    formula mudam a cada coleta: o coletor grava valor onde havia formula, os
    consolidadores criam formula onde nao havia. O indice passa a apontar para
    celula que nao e mais formula, e o Excel declara o arquivo corrompido:

        Removed Records: Formula from /xl/calcChain.xml part

    Com Notify=False, esse dialogo vira "Open method of Workbooks class failed" e
    o passe do Excel perde a planilha inteira. Foi o que travou FrangoUS,
    CarRental e WEG em todo ciclo entre 19 e 20/08.

    Apagar e seguro: o Excel reconstroi o calcChain sozinho no primeiro
    recalculo. O BeefBR.xlsm nunca teve essa parte e sempre abriu normalmente.
    """
    import re
    import shutil
    import tempfile
    import zipfile

    with zipfile.ZipFile(path, "r") as zin:
        if "xl/calcChain.xml" not in zin.namelist():
            return  # nada a fazer
        fd, tmp = tempfile.mkstemp(suffix=".xlsm")
        os.close(fd)
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "xl/calcChain.xml":
                    continue
                data = zin.read(item.filename)
                # A parte sai do pacote; as duas referencias a ela tambem, senao
                # o Excel reclama de relacionamento apontando para o vazio.
                if item.filename == "[Content_Types].xml":
                    data = re.sub(rb"<Override[^>]*calcChain\.xml[^>]*/>", b"", data)
                elif item.filename.endswith(".rels"):
                    data = re.sub(rb"<Relationship[^>]*calcChain\.xml[^>]*/>", b"", data)
                zout.writestr(item, data)
    shutil.move(tmp, path)
    print("[Drive] calcChain removido (o Excel reconstroi).", flush=True)


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

    # Opera direto na tag <f ...>, sem casar a celula <c> em volta: `[^>]*` nao
    # atravessa tags, entao o custo e' linear. A versao anterior casava
    # `<c ...>.*?</c>` com DOTALL e travava em planilhas salvas pelo openpyxl,
    # que grava celula vazia auto-fechada (<c r="A1" s="5"/>): sem `</c>`
    # proprio, cada uma disparava uma varredura do documento inteiro.
    #
    # Dispensar o contexto da celula nao muda o resultado: a regra antiga era
    # "remove se ref == referencia da celula OU se ref e' uma celula unica", e
    # a referencia da celula e' sempre uma celula unica — logo a primeira
    # condicao ja' esta contida na segunda.
    formula_re = re.compile(r"<f\b([^>]*)>")
    ref_unica_re = re.compile(r"^\$?[A-Za-z]{1,3}\$?\d+$")

    def normalize_formula(match: re.Match) -> str:
        attrs = match.group(1)
        if 't="array"' not in attrs:
            return match.group(0)

        ref_match = re.search(r'\bref="([^"]+)"', attrs)
        if not ref_match or not ref_unica_re.match(ref_match.group(1)):
            return match.group(0)

        attrs = re.sub(r'\s+t="array"', "", attrs)
        attrs = re.sub(r'\s+ref="[^"]+"', "", attrs)
        return f"<f{attrs}>"

    mudou = False
    with zipfile.ZipFile(path, "r") as zin:
        fd, tmp = tempfile.mkstemp(suffix=".xlsm")
        os.close(fd)
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if re.match(r"xl/worksheets/sheet\d+\.xml$", item.filename):
                    txt = data.decode("utf-8")
                    # Sem formula matricial na aba nao ha' o que normalizar.
                    novo = formula_re.sub(normalize_formula, txt) if 't="array"' in txt else txt
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
    shared_path = os.environ.get(SHARED_PATH_ENV)
    shared_file_id = os.environ.get(SHARED_FILE_ID_ENV)
    if (
        shared_path
        and shared_file_id == file_id
        and Path(shared_path).resolve() == Path(local_path).resolve()
    ):
        print("[Drive] Upload adiado para o fim do pipeline.", flush=True)
        return {"id": file_id, "deferred": True}

    print("[Drive] Preparando upload da planilha...", flush=True)
    t0 = time.monotonic()
    _fix_comment_paths(local_path)
    _fix_single_cell_array_formulas(local_path)
    _drop_calc_chain(local_path)
    # Upload retomavel: se a conexao TLS cair no meio do envio, o cliente
    # repete somente o chunk interrompido em vez de derrubar o workflow.
    media = MediaFileUpload(
        local_path,
        mimetype=mime,
        chunksize=DRIVE_UPLOAD_CHUNK_SIZE,
        resumable=True,
    )
    request = (
        drive()
        .files()
        .update(
            fileId=file_id,
            media_body=media,
            fields="id,name,size,modifiedTime,md5Checksum",
            supportsAllDrives=True,
        )
    )

    tamanho_mb = os.path.getsize(local_path) / (1024 * 1024)
    print(
        f"[Drive] Enviando {tamanho_mb:.1f} MB em pedacos de "
        f"{DRIVE_UPLOAD_CHUNK_SIZE // (1024 * 1024)} MB...",
        flush=True,
    )

    res = None
    ultimo_progresso = -1
    while res is None:
        t_chunk = time.monotonic()
        status, res = request.next_chunk(num_retries=DRIVE_NUM_RETRIES)
        # Log por pedaco (e nao so' quando a % muda): num upload lento e' isso
        # que diferencia "travado" de "andando devagar".
        if status:
            progresso = int(status.progress() * 100)
            if progresso != ultimo_progresso:
                print(
                    f"[Drive] Upload {progresso}% "
                    f"({time.monotonic() - t_chunk:.1f}s neste pedaco, "
                    f"{time.monotonic() - t0:.1f}s no total).",
                    flush=True,
                )
                ultimo_progresso = progresso

    print(f"[Drive] Upload concluido em {time.monotonic() - t0:.1f}s.", flush=True)
    return res
