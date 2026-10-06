"""Leitura do Analítico e envio em lotes (API falsa; nada sai da máquina)."""
import datetime as dt
import os
import re
import time
import zipfile

import openpyxl
import pytest

try:
    from utils import fluxo_desembolso as fd
except ImportError:
    import fluxo_desembolso as fd


class ApiFalsa:
    def __init__(self, inicio_dados=1, fim=None):
        self.inicio_dados = inicio_dados
        self.fim = {"obras_reconhecidas": ["A"], "obras_nao_reconhecidas": [], "obras_em_conflito": []} \
            if fim is None else fim
        self.chamadas = []

    def conferir_acesso(self):
        self.chamadas.append(("verificar",))

    def iniciar(self, sienge, arquivo_nome, primeiras_linhas):
        self.chamadas.append(("iniciar", sienge, arquivo_nome, primeiras_linhas))
        return {"envio_id": "e-1", "inicio_dados": self.inicio_dados}

    def lote(self, envio_id, seq, linhas):
        self.chamadas.append(("lote", envio_id, seq, linhas))
        return {"seq": seq, "lancamentos": len(linhas), "ignoradas": 1}

    def concluir(self, envio_id, total_lotes):
        self.chamadas.append(("concluir", envio_id, total_lotes))
        return self.fim


def xlsx(tmp_path, linhas, nome="V.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    for linha in linhas:
        ws.append(linha)
    wb.create_sheet("Outra").append(["não lida"])
    caminho = tmp_path / nome
    wb.save(caminho)
    return str(caminho)


@pytest.mark.parametrize("valor, esperado", [
    (None, ""),
    ("", ""),
    ("  texto com espaço ", "  texto com espaço "),
    ("01/02/2026", "01/02/2026"),
    (dt.datetime(2026, 3, 5, 14, 30), "2026-03-05"),
    (dt.datetime(2026, 3, 5, 23, 59, 59), "2026-03-05"),
    (dt.date(2001, 12, 31), "2001-12-31"),
    (10, 10),
    (-1234.56, -1234.56),
    (True, "True"),
])
def test_converter_celula(valor, esperado):
    convertido = fd.converter_celula(valor)
    assert convertido == esperado and type(convertido) is type(esperado)


def test_ler_linhas_primeira_aba_converte_e_descarta_vazias(tmp_path):
    caminho = xlsx(tmp_path, [
        ["Relatório"], [None, None], ["Obra", "Data", "Valor"],
        ["1 - X", dt.datetime(2026, 1, 2), 10.5], [None, None, None], ["2 - Y", None, 3],
    ])
    assert fd.ler_linhas(caminho) == [
        ["Relatório"], ["Obra", "Data", "Valor"], ["1 - X", "2026-01-02", 10.5], ["2 - Y", "", 3],
    ]


def test_ler_linhas_ignora_dimensao_errada_declarada(tmp_path):
    caminho = xlsx(tmp_path, [["Obra", "Valor"], ["1 - X", 1], ["2 - Y", 2]])
    # O Sienge grava <dimension ref="A1"/>; sem reset_dimensions o openpyxl leria só A1.
    origem = zipfile.ZipFile(caminho)
    partes = {n: origem.read(n) for n in origem.namelist()}
    origem.close()
    planilha = "xl/worksheets/sheet1.xml"
    partes[planilha] = re.sub(rb'<dimension ref="[^"]*"\s*/>', b'<dimension ref="A1"/>', partes[planilha])
    assert b'<dimension ref="A1"/>' in partes[planilha]
    with zipfile.ZipFile(caminho, "w") as destino:
        for nome, conteudo in partes.items():
            destino.writestr(nome, conteudo)
    assert fd.ler_linhas(caminho) == [["Obra", "Valor"], ["1 - X", 1], ["2 - Y", 2]]


def test_envia_iniciar_lotes_e_concluir_na_ordem(tmp_path):
    caminho = xlsx(tmp_path, [["Preâmbulo"], [None], ["Obra", "Valor"], ["1 - X", 1], ["2 - Y", 2]])
    api = ApiFalsa(inicio_dados=2)  # índice na lista SEM a linha vazia
    log = []
    resumo = fd.enviar_analitico(caminho, "habitat", api, log=log.append)
    assert api.chamadas == [
        ("verificar",),
        ("iniciar", "habitat", "V.xlsx", [["Preâmbulo"], ["Obra", "Valor"], ["1 - X", 1], ["2 - Y", 2]]),
        ("lote", "e-1", 1, [["1 - X", 1], ["2 - Y", 2]]),
        ("concluir", "e-1", 1),
    ]
    assert resumo["lotes"] == 1 and resumo["lancamentos"] == 2 and resumo["ignoradas"] == 1
    assert resumo["envio_id"] == "e-1" and resumo["linhas"] == 4
    assert any("V.xlsx" in m for m in log)


def test_primeiras_linhas_limitadas_a_40(tmp_path, monkeypatch):
    caminho = xlsx(tmp_path, [["x"]])
    monkeypatch.setattr(fd, "ler_linhas", lambda c: [[i] for i in range(100)])
    api = ApiFalsa(inicio_dados=1)
    fd.enviar_analitico(caminho, "top", api, log=lambda m: None)
    assert api.chamadas[1][3] == [[i] for i in range(40)]


@pytest.mark.parametrize("n_dados, esperados", [(5000, [5000]), (5001, [5000, 1]), (12345, [5000, 5000, 2345])])
def test_fatia_em_lotes_de_5000(tmp_path, monkeypatch, n_dados, esperados):
    caminho = xlsx(tmp_path, [["x"]])
    monkeypatch.setattr(fd, "ler_linhas", lambda c: [["Obra"]] + [[i] for i in range(n_dados)])
    api = ApiFalsa(inicio_dados=1)
    resumo = fd.enviar_analitico(caminho, "top", api, log=lambda m: None)
    lotes = [c for c in api.chamadas if c[0] == "lote"]
    assert [len(c[3]) for c in lotes] == esperados
    assert [c[2] for c in lotes] == list(range(1, len(esperados) + 1))
    assert lotes[0][3][0] == [0] and lotes[-1][3][-1] == [n_dados - 1]
    assert api.chamadas[-1] == ("concluir", "e-1", len(esperados))
    assert resumo["lotes"] == len(esperados)


def test_sem_dados_nao_conclui(tmp_path):
    caminho = xlsx(tmp_path, [["Obra", "Valor"]])
    api = ApiFalsa(inicio_dados=1)
    with pytest.raises(fd.PlanilhaSemDados):
        fd.enviar_analitico(caminho, "top", api, log=lambda m: None)
    assert [c[0] for c in api.chamadas] == ["verificar", "iniciar"]


def test_arquivo_inexistente(tmp_path):
    api = ApiFalsa()
    with pytest.raises(FileNotFoundError, match="nao_existe.xlsx"):
        fd.enviar_analitico(str(tmp_path / "nao_existe.xlsx"), "top", api, log=lambda m: None)
    assert api.chamadas == []


def test_arquivo_velho_recusado_sem_chamar_o_gio(tmp_path):
    caminho = xlsx(tmp_path, [["Obra"], ["1 - X"]])
    velho = time.time() - 21 * 3600
    os.utime(caminho, (velho, velho))
    api = ApiFalsa()
    with pytest.raises(fd.ArquivoVelho, match="nada enviado"):
        fd.enviar_analitico(caminho, "top", api, log=lambda m: None)
    assert api.chamadas == []


def test_arquivo_de_19h_passa_e_idade_none_desliga(tmp_path):
    caminho = xlsx(tmp_path, [["Obra"], ["1 - X"]])
    t = time.time() - 19 * 3600
    os.utime(caminho, (t, t))
    fd.enviar_analitico(caminho, "top", ApiFalsa(), log=lambda m: None)
    t = time.time() - 90 * 24 * 3600
    os.utime(caminho, (t, t))
    fd.enviar_analitico(caminho, "top", ApiFalsa(), idade_max_h=None, log=lambda m: None)


def test_xlsx_corrompido_falha_antes_de_iniciar(tmp_path):
    caminho = tmp_path / "V.xlsx"
    caminho.write_bytes(b"PK\x03\x04 download pela metade")
    api = ApiFalsa()
    with pytest.raises(zipfile.BadZipFile):
        fd.enviar_analitico(str(caminho), "top", api, log=lambda m: None)
    assert "iniciar" not in [c[0] for c in api.chamadas]


def test_avisa_obras_pendentes_sem_falhar(tmp_path):
    caminho = xlsx(tmp_path, [["Obra"], ["1 - X"]])
    fim = {"obras_reconhecidas": [], "obras_nao_reconhecidas": ["33281 - BRISAS"], "obras_em_conflito": ["9 - Z"]}
    log = []
    fd.enviar_analitico(caminho, "habitat", ApiFalsa(fim=fim), log=log.append)
    texto = "\n".join(log)
    assert "33281 - BRISAS" in texto and "9 - Z" in texto and "Configurações" in texto


def test_avisa_concluido_por_reenvio(tmp_path):
    caminho = xlsx(tmp_path, [["Obra"], ["1 - X"]])
    log = []
    resumo = fd.enviar_analitico(caminho, "top", ApiFalsa(fim={"concluido_por_reenvio": True}), log=log.append)
    assert resumo["concluido_por_reenvio"] is True
    assert any("já tinha gravado" in m for m in log)


def test_chaves_do_robo_prevalecem_sobre_as_do_concluir(tmp_path):
    caminho = xlsx(tmp_path, [["Obra"], ["1 - X"]])
    fim = {"linhas": 999, "lotes": 999, "obras_reconhecidas": []}
    resumo = fd.enviar_analitico(caminho, "top", ApiFalsa(fim=fim), log=lambda m: None)
    assert resumo["linhas"] == 2 and resumo["lotes"] == 1
