"""Cliente da API de ingestão do GIO (Edge Function `ingest-fluxo-desembolso`).

Ações: `verificar` (confere URL e chave), `iniciar`, `lote` e `concluir` — contrato em
docs/integracoes/ingest-fluxo-desembolso.md, no repositório do GIO. A chave vai no header
`x-api-key` e nunca aparece em log nem em mensagem de erro. Arquivo idêntico nos 5 robôs Sienge.
"""
from __future__ import annotations

import os
import time
from typing import Callable, Mapping
from urllib.parse import urlsplit

import requests

TIMEOUT = 120
TENTATIVAS = 3  # a Central não refaz execução em que só este módulo falhou
PAUSA_ENTRE_TENTATIVAS = 30
HOSTS_LOCAIS = ("127.0.0.1", "localhost")
# Trecho comum às duas mensagens do servidor: "Envio não encontrado (expirado ou já concluído)." e
# "Envio sem lançamentos, expirado ou já concluído."
JA_CONCLUIDO = "expirado ou já concluído"


class ErroApi(Exception):
    """Falha da chamada inteira (rede, 4xx, 5xx). Derruba o módulo em execução."""

    def __init__(self, mensagem: str, status: int | None = None, do_servidor: str = ""):
        super().__init__(mensagem)
        self.status = status
        self.do_servidor = do_servidor  # mensagem que o servidor mandou, para o diagnóstico


def _concluir_ja_gravou(status: int, do_servidor: str) -> bool:
    """Contrato: 400 'expirado ou já concluído' depois de 500/timeout no `concluir` = a 1ª chamada gravou."""
    return status == 400 and JA_CONCLUIDO in do_servidor


class GioApi:
    def __init__(self, url: str | None, chave: str | None, sessao=None, pausa=time.sleep):
        url = (url or "").strip()
        chave = (chave or "").strip()
        if not url or not chave:
            raise ErroApi("Defina GIO_INGEST_URL e GIO_INGEST_KEY no .env.")
        partes = urlsplit(url)
        if partes.scheme != "https" and partes.hostname not in HOSTS_LOCAIS:
            raise ErroApi("GIO_INGEST_URL deve começar com https://.")
        self.url = url
        self._chave = chave
        self._sessao = sessao or requests.Session()
        self._pausa = pausa

    @classmethod
    def do_ambiente(cls, env: Mapping[str, str] | None = None, **opcoes) -> "GioApi":
        env = os.environ if env is None else env
        return cls(env.get("GIO_INGEST_URL"), env.get("GIO_INGEST_KEY"), **opcoes)

    def conferir_acesso(self) -> None:
        """Confere URL e chave sem gravar nada no GIO."""
        try:
            dados = self._chamar({"acao": "verificar"})
        except ErroApi as erro:
            if erro.status == 401:
                raise ErroApi(f"O GIO respondeu 401 ({erro.do_servidor}). Confira GIO_INGEST_KEY; se a mensagem "
                              'fala de "authorization", a função foi publicada sem verify_jwt = false.', 401) from None
            if erro.status == 404:
                raise ErroApi("GIO_INGEST_URL não existe (404). A função foi publicada?", 404) from None
            raise
        if dados.get("ok") is not True:
            raise ErroApi("Resposta inesperada de 'verificar'; confira GIO_INGEST_URL.")

    def iniciar(self, sienge: str, arquivo_nome: str, primeiras_linhas: list) -> dict:
        """Abre o envio; devolve `envio_id` e `inicio_dados` (índice base 0 da 1ª linha de dados)."""
        dados = self._chamar({"acao": "iniciar", "sienge": sienge, "arquivo_nome": arquivo_nome,
                              "primeiras_linhas": primeiras_linhas})
        envio_id = dados.get("envio_id")
        inicio = dados.get("inicio_dados")
        if not isinstance(envio_id, str) or not envio_id:
            raise ErroApi("Resposta de 'iniciar' sem 'envio_id'.")
        if type(inicio) is not int or inicio < 0:
            raise ErroApi("Resposta de 'iniciar' sem 'inicio_dados' válido.")
        return dados

    def lote(self, envio_id: str, seq: int, linhas: list) -> dict:
        """Envia um lote (≤ 5.000 linhas). Reenviar o mesmo `seq` substitui o anterior no servidor."""
        return self._chamar({"acao": "lote", "envio_id": envio_id, "seq": seq, "linhas": linhas})

    def concluir(self, envio_id: str, total_lotes: int) -> dict:
        """Fecha o envio e grava. `{"concluido_por_reenvio": True}` = a tentativa anterior já tinha gravado."""
        return self._chamar({"acao": "concluir", "envio_id": envio_id, "total_lotes": total_lotes},
                            sucesso_apos_falha=_concluir_ja_gravou)

    def _chamar(self, corpo: dict, sucesso_apos_falha: Callable[[int, str], bool] | None = None) -> dict:
        ultimo = ""
        status = None
        houve_falha = False  # 500 ou sem resposta numa tentativa anterior desta mesma chamada
        for tentativa in range(1, TENTATIVAS + 1):
            status = None
            try:
                resposta = self._sessao.post(
                    self.url, json=corpo, headers={"x-api-key": self._chave}, timeout=TIMEOUT,
                    allow_redirects=False)  # seguir levaria a chave a outro host ou a http
            except requests.RequestException as erro:
                ultimo = f"sem resposta do GIO ({type(erro).__name__})"
                houve_falha = True
            else:
                if resposta.status_code == 200:
                    try:
                        dados = resposta.json()
                    except ValueError:
                        dados = None
                    if isinstance(dados, dict):
                        return dados
                    raise ErroApi("GIO respondeu 200 com corpo que não é JSON.")
                status = resposta.status_code
                if 300 <= status < 400:
                    raise ErroApi(
                        f"GIO respondeu {status}: redirecionamento recusado (confira GIO_INGEST_URL).", status)
                do_servidor = self._mensagem(resposta)
                ultimo = f"GIO respondeu {status}: {do_servidor}"
                if status < 500:
                    if houve_falha and sucesso_apos_falha and sucesso_apos_falha(status, do_servidor):
                        return {"concluido_por_reenvio": True}
                    raise ErroApi(ultimo, status, do_servidor)
                houve_falha = True
            if tentativa < TENTATIVAS:
                self._pausa(PAUSA_ENTRE_TENTATIVAS)
        raise ErroApi(ultimo, status)

    @staticmethod
    def _mensagem(resposta) -> str:
        try:
            dados = resposta.json()
        except ValueError:
            dados = None
        if isinstance(dados, dict) and dados.get("error"):
            return " ".join(str(dados["error"]).split())[:200] or "sem mensagem"
        return " ".join((resposta.text or "").split())[:200] or "sem mensagem"
