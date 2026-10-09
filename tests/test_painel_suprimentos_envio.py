"""Envio do Painel de Suprimentos em lotes (API falsa; nada sai da máquina)."""
import os
import time

import openpyxl
import pytest

try:
    from utils import painel_suprimentos_envio as ps
except ImportError:
    import painel_suprimentos_envio as ps

CAB = ["N° do Pedido", "Obra", "Data do pedido", "Saldo"]


class ApiFalsa:
    def __init__(self, inicio_dados=1, fim=None):
        self.inicio_dados = inicio_dados
        self.fim = {"pedidos": 2, "itens": 3, "obras_reconhecidas": ["A"], "obras_nao_reconhecidas": ["Z"]} \
            if fim is None else fim
        self.chamadas = []

    def conferir_acesso(self):
        self.chamadas.append(("verificar",))

    def iniciar(self, sienge, arquivo_nome, primeiras_linhas):
        self.chamadas.append(("iniciar", sienge, arquivo_nome, len(primeiras_linhas)))
        return {"envio_id": "e-1", "inicio_dados": self.inicio_dados}

    def lote(self, envio_id, seq, linhas):
        self.chamadas.append(("lote", seq, len(linhas)))
        return {"seq": seq, "linhas": len(linhas)}

    def concluir(self, envio_id, total_lotes):
        self.chamadas.append(("concluir", total_lotes))
        return self.fim


@pytest.fixture(autouse=True)
def minimo_baixo(monkeypatch):
    # Planilhas de teste são pequenas; o mínimo real tem teste próprio.
    monkeypatch.setattr(ps, "MIN_LINHAS_DADOS", 1)


def xlsx(tmp_path, linhas, nome="PAINEL DE SUPRIMENTOS - TOP.xlsx"):
    wb = openpyxl.Workbook()
    for linha in linhas:
        wb.active.append(linha)
    caminho = tmp_path / nome
    wb.save(caminho)
    return str(caminho)


def test_envia_em_lotes_a_partir_de_inicio_dados(tmp_path, monkeypatch):
    monkeypatch.setattr(ps, "TAMANHO_LOTE", 2)
    caminho = xlsx(tmp_path, [CAB] + [[str(i), "A", "2026-09-01", 1] for i in range(5)])
    api = ApiFalsa()
    logs = []
    r = ps.enviar_painel(caminho, "top", api, log=logs.append)
    assert api.chamadas == [("verificar",), ("iniciar", "top", "PAINEL DE SUPRIMENTOS - TOP.xlsx", 6),
                            ("lote", 1, 2), ("lote", 2, 2), ("lote", 3, 1), ("concluir", 3)]
    assert r["lotes"] == 3 and r["pedidos"] == 2
    assert any("Z" in l and "Mapear Obras" in l for l in logs)


def test_arquivo_velho_nao_envia(tmp_path):
    caminho = xlsx(tmp_path, [CAB, ["1", "A", "", 1]])
    api = ApiFalsa()
    with pytest.raises(ps.ArquivoVelho):
        ps.enviar_painel(caminho, "top", api, agora=lambda: os.path.getmtime(caminho) + 21 * 3600)
    assert api.chamadas == []


def test_arquivo_ausente(tmp_path):
    with pytest.raises(FileNotFoundError):
        ps.enviar_painel(str(tmp_path / "nao.xlsx"), "top", ApiFalsa())


def test_sem_linhas_de_dados(tmp_path):
    caminho = xlsx(tmp_path, [CAB])
    with pytest.raises(ps.PlanilhaSemDados):
        ps.enviar_painel(caminho, "top", ApiFalsa())


def test_concluido_por_reenvio_avisa(tmp_path):
    caminho = xlsx(tmp_path, [CAB, ["1", "A", "", 1]])
    logs = []
    ps.enviar_painel(caminho, "top", ApiFalsa(fim={"concluido_por_reenvio": True}), log=logs.append)
    assert any("AVISO" in l for l in logs)


def test_api_do_ambiente_exige_as_duas_variaveis():
    with pytest.raises(Exception, match="GIO_SUPRIMENTOS_URL"):
        ps.api_do_ambiente({"GIO_SUPRIMENTOS_URL": "https://x"})


def test_planilha_truncada_abaixo_do_minimo_nao_envia(tmp_path, monkeypatch):
    monkeypatch.setattr(ps, "MIN_LINHAS_DADOS", 100)
    caminho = xlsx(tmp_path, [CAB] + [[str(i), "A", "", 1] for i in range(14)])
    api = ApiFalsa()
    with pytest.raises(ps.PlanilhaTruncada, match="14 linhas"):
        ps.enviar_painel(caminho, "top", api)
    assert api.chamadas == []  # nem o verificar: nada sai da máquina


def test_minimo_padrao_e_100(monkeypatch):
    monkeypatch.undo()  # desfaz o minimo_baixo
    assert ps.MIN_LINHAS_DADOS == 100
