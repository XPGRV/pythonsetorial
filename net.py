"""Deixa o `requests` resiliente a instabilidade das fontes (timeouts e erros
transitorios 5xx), que e comum nos servidores .gov.br.

Basta importar este modulo uma vez no processo (o `drive_io` ja faz isso) que
TODAS as chamadas `requests.get/post` e `requests.Session()` passam a repetir
automaticamente em caso de falha de conexao/leitura, com backoff.

Nao muda o codigo dos coletores.
"""

import requests
from requests.adapters import HTTPAdapter

try:
    from urllib3.util.retry import Retry
except ImportError:  # fallback p/ versoes antigas
    from requests.packages.urllib3.util.retry import Retry

# 5 tentativas, backoff 3s -> 0/3/6/12/24s (limitado). Retry em conexao, leitura
# e status 5xx/429. allowed_methods=None inclui POST (as consultas SECEX sao POST
# idempotentes, seguras de repetir).
_RETRY = Retry(
    total=5,
    connect=5,
    read=5,
    backoff_factor=3,
    status_forcelist=(429, 500, 502, 503, 504),
    allowed_methods=None,
    raise_on_status=False,
)


def _mount(session: requests.Session) -> requests.Session:
    adapter = HTTPAdapter(max_retries=_RETRY)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    return session


# Aplica o retry em toda Session nova. `requests.get/post` criam uma Session
# interna, entao isso cobre tambem as chamadas diretas.
if not getattr(requests.Session, "_retry_patched", False):
    _orig_init = requests.Session.__init__

    def _init(self, *args, **kwargs):
        _orig_init(self, *args, **kwargs)
        _mount(self)

    requests.Session.__init__ = _init
    requests.Session._retry_patched = True
