"""Envio do Analítico de Apropriação (VENCIMENTO) ao GIO, em lotes.

Só lê a planilha e converte células; parse, validação e casamento de obras ficam no servidor
(Edge Function `ingest-fluxo-desembolso`, contrato em docs/integracoes/ingest-fluxo-desembolso.md
no repositório do GIO). Arquivo idêntico nos 5 robôs Sienge.
"""
from __future__ import annotations

import datetime as dt
import os
import time

import openpyxl

LINHAS_INICIO = 40     # o `iniciar` recebe o preâmbulo + cabeçalho
TAMANHO_LOTE = 5000    # máximo aceito por `lote`
IDADE_MAX_H = 20       # mais velho que isso = o Analítico do dia não foi gerado


class ArquivoVelho(Exception):
    """O arquivo em relatorios/ não é do dia: enviar regrediria o Acompanhamento sem ninguém ver."""


class PlanilhaSemDados(Exception):
    """Nenhuma linha depois do cabeçalho."""


def converter_celula(v):
    """data → 'YYYY-MM-DD' (partes da própria data, sem fuso); número → número; vazio → ''; resto → texto."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, (dt.datetime, dt.date)):
        return f"{v.year:04d}-{v.month:02d}-{v.day:02d}"
    if isinstance(v, (int, float)):
        return v
    return str(v)


def ler_linhas(caminho: str) -> list:
    """Primeira aba inteira, células convertidas, sem as linhas totalmente vazias."""
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    try:
        ws = wb.worksheets[0]
        ws.reset_dimensions()  # o Sienge declara a dimensão errada; sem isto o openpyxl corta a planilha
        linhas = []
        for linha in ws.iter_rows(values_only=True):
            convertida = [converter_celula(v) for v in linha]
            if any(c != "" for c in convertida):
                linhas.append(convertida)
        return linhas
    finally:
        wb.close()


def enviar_analitico(caminho: str, sienge: str, api, idade_max_h: float | None = IDADE_MAX_H,
                     agora=time.time, log=print) -> dict:
    """verificar → iniciar → lotes de 5.000 → concluir. Levanta exceção em qualquer falha (o módulo falha).

    `idade_max_h=None` desliga a checagem de idade (só teste ou carga manual).
    """
    if not os.path.isfile(caminho):
        raise FileNotFoundError(f"Analítico VENCIMENTO não encontrado: {caminho}")
    mtime = os.path.getmtime(caminho)
    quando = dt.datetime.fromtimestamp(mtime).strftime("%d/%m/%Y %H:%M")
    if idade_max_h is not None and agora() - mtime > idade_max_h * 3600:
        raise ArquivoVelho(f"O Analítico em {caminho} é de {quando} (mais de {idade_max_h} h): o Analítico do dia "
                           "não foi gerado; nada enviado ao GIO.")

    api.conferir_acesso()
    linhas = ler_linhas(caminho)
    nome = os.path.basename(caminho)
    log(f"Fluxo de Desembolso: {nome} ({quando}), {len(linhas)} linhas lidas; Sienge '{sienge}'.")

    inicio = api.iniciar(sienge, nome, linhas[:LINHAS_INICIO])
    envio_id = inicio["envio_id"]
    dados = linhas[inicio["inicio_dados"]:]
    if not dados:
        # O envio fica aberto no servidor, que o apaga depois de 2 dias.
        raise PlanilhaSemDados("O Analítico não tem linhas de dados depois do cabeçalho; nada gravado no GIO.")

    lancamentos = ignoradas = total_lotes = 0
    for i in range(0, len(dados), TAMANHO_LOTE):
        total_lotes += 1
        resposta = api.lote(envio_id, total_lotes, dados[i:i + TAMANHO_LOTE])
        lancamentos += int(resposta.get("lancamentos") or 0)
        ignoradas += int(resposta.get("ignoradas") or 0)

    fim = api.concluir(envio_id, total_lotes)
    if fim.get("concluido_por_reenvio"):
        log("AVISO: o 'concluir' não respondeu de primeira e o reenvio achou o envio já fechado: a primeira "
            "chamada já tinha gravado. Conferir o Acompanhamento.")
    log(f"Fluxo de Desembolso: {total_lotes} lote(s), {lancamentos} lançamentos, {ignoradas} linhas ignoradas, "
        f"{len(fim.get('obras_reconhecidas') or [])} obras reconhecidas.")
    for chave, rotulo in (("obras_nao_reconhecidas", "obra(s) não reconhecida(s)"),
                          ("obras_em_conflito", "obra(s) em conflito com outro Sienge")):
        obras = fim.get(chave) or []
        if obras:
            log(f"AVISO: {len(obras)} {rotulo} — resolver na aba Configurações do Acompanhamento: "
                + ", ".join(str(o) for o in obras))
    # As chaves do robô vêm depois: o `concluir` também devolve `linhas`/`lotes` (contagens do servidor).
    return {**fim, "envio_id": envio_id, "linhas": len(linhas), "lotes": total_lotes,
            "lancamentos": lancamentos, "ignoradas": ignoradas}
