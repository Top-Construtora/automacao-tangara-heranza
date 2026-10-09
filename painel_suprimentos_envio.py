"""Envio do PAINEL DE SUPRIMENTOS ao GIO (Painel de Suprimentos), em lotes.

Só lê a planilha e converte células (reaproveita `ler_linhas` do envio do Fluxo de Desembolso);
agregação item → pedido e casamento de obras ficam no servidor (Edge Function
`ingest-painel-suprimentos`, contrato em docs/integracoes/ingest-painel-suprimentos.md no
repositório do GIO). Arquivo idêntico nos 5 robôs Sienge.
"""
from __future__ import annotations

import datetime as dt
import os
import time
from typing import Mapping

try:
    from utils.fluxo_desembolso import ler_linhas
    from utils.gio_api import ErroApi, GioApi
except ImportError:  # monolitos: arquivos na raiz
    from fluxo_desembolso import ler_linhas
    from gio_api import ErroApi, GioApi

LINHAS_INICIO = 40     # o `iniciar` recebe as primeiras linhas (cabeçalho incluso)
TAMANHO_LOTE = 5000    # máximo aceito por `lote`
IDADE_MAX_H = 20       # mais velho que isso = o Painel do dia não foi exportado
MIN_LINHAS_DADOS = 100  # menos que isso = export truncado (grade do Sienge que não carregou)


class ArquivoVelho(Exception):
    """O arquivo em relatorios/ não é do dia: enviar regrediria o Painel sem ninguém ver."""


class PlanilhaSemDados(Exception):
    """Nenhuma linha depois do cabeçalho."""


class PlanilhaTruncada(PlanilhaSemDados):
    """Poucas linhas: o Painel de Compras não carregou a grade inteira antes do export."""


def api_do_ambiente(env: Mapping[str, str] | None = None) -> GioApi:
    """GioApi apontada para a função do Painel (URL e chave próprias, separadas das do Fluxo)."""
    env = os.environ if env is None else env
    url = (env.get("GIO_SUPRIMENTOS_URL") or "").strip()
    chave = (env.get("GIO_SUPRIMENTOS_KEY") or "").strip()
    if not url or not chave:
        raise ErroApi("Defina GIO_SUPRIMENTOS_URL e GIO_SUPRIMENTOS_KEY no .env.")
    return GioApi(url, chave)


def enviar_painel(caminho: str, sienge: str, api, idade_max_h: float | None = IDADE_MAX_H,
                  agora=time.time, log=print) -> dict:
    """verificar → iniciar → lotes de 5.000 → concluir. Levanta exceção em qualquer falha (o módulo falha)."""
    if not os.path.isfile(caminho):
        raise FileNotFoundError(f"Painel de Suprimentos não encontrado: {caminho}")
    mtime = os.path.getmtime(caminho)
    quando = dt.datetime.fromtimestamp(mtime).strftime("%d/%m/%Y %H:%M")
    if idade_max_h is not None and agora() - mtime > idade_max_h * 3600:
        raise ArquivoVelho(f"O Painel em {caminho} é de {quando} (mais de {idade_max_h} h): o Painel do dia "
                           "não foi exportado; nada enviado ao GIO.")

    linhas = ler_linhas(caminho)
    nome = os.path.basename(caminho)
    # Antes de chamar o GIO: o export pode sair "com sucesso" só com o cabeçalho ou parte da grade.
    if len(linhas) - 1 < MIN_LINHAS_DADOS:
        raise PlanilhaTruncada(f"O Painel em {caminho} ({quando}) tem só {max(len(linhas) - 1, 0)} linhas de dados "
                               f"(mínimo {MIN_LINHAS_DADOS}): export truncado; nada enviado ao GIO.")
    api.conferir_acesso()
    log(f"Painel de Suprimentos: {nome} ({quando}), {len(linhas)} linhas lidas; Sienge '{sienge}'.")

    inicio = api.iniciar(sienge, nome, linhas[:LINHAS_INICIO])
    envio_id = inicio["envio_id"]
    dados = linhas[inicio["inicio_dados"]:]
    if not dados:
        # O envio fica aberto no servidor, que o apaga depois de 24 h.
        raise PlanilhaSemDados("O Painel não tem linhas de dados depois do cabeçalho; nada gravado no GIO.")

    total_lotes = 0
    for i in range(0, len(dados), TAMANHO_LOTE):
        total_lotes += 1
        api.lote(envio_id, total_lotes, dados[i:i + TAMANHO_LOTE])

    fim = api.concluir(envio_id, total_lotes)
    if fim.get("concluido_por_reenvio"):
        log("AVISO: o 'concluir' não respondeu de primeira e o reenvio achou o envio já fechado: a primeira "
            "chamada já tinha gravado. Conferir o Painel de Suprimentos.")
        log(f"Painel de Suprimentos: {total_lotes} lote(s).")
    else:
        log(f"Painel de Suprimentos: {total_lotes} lote(s), {fim.get('pedidos')} pedidos, {fim.get('itens')} itens, "
            f"{len(fim.get('obras_reconhecidas') or [])} obras reconhecidas.")
    nao = fim.get("obras_nao_reconhecidas") or []
    if nao:
        log(f"AVISO: {len(nao)} obra(s) não reconhecida(s) — resolver em Mapear Obras no Painel de Suprimentos: "
            + ", ".join(str(o) for o in nao))
    return {**fim, "envio_id": envio_id, "linhas": len(linhas), "lotes": total_lotes}
