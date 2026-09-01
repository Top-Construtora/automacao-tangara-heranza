"""Política de exportação do MuiDataGrid (Cadastro de Contratos, Sienge 9.0.4).

O 'Gerar Relatório' exporta só as linhas materializadas no grid. Antes de
exportar é preciso: trocar a paginação para 'Todas/Todos', esperar as linhas
chegarem e, se não chegarem, tentar de novo — e por fim FALHAR em vez de
entregar ao BI um arquivo com uma página só (26 linhas).
"""
import pytest

import main


class RelogioFalso:
    def __init__(self):
        self.agora = 1_000.0

    def time(self):
        return self.agora

    def sleep(self, segundos):
        self.agora += segundos


class DriverFalso:
    """execute_script devolve o status do grid em função do relógio."""

    def __init__(self, status_fn):
        self.status_fn = status_fn

    def execute_script(self, js, *args):
        return self.status_fn()


PRIMEIRA_PAGINA = {"carregadas": 26, "total": 127, "carregando": False}
COMPLETO = {"carregadas": 129, "total": 127, "carregando": False}


@pytest.fixture
def relogio(monkeypatch):
    r = RelogioFalso()
    monkeypatch.setattr(main, "time", r)
    return r


@pytest.fixture
def paginacao(monkeypatch):
    """Substitui a interação Selenium com o select de paginação por um registro
    de chamadas; cada chamada aplica o próximo 'efeito' configurado."""
    estado = {"chamadas": 0, "efeitos": []}

    def falsa(driver, wait, *a, **k):
        estado["chamadas"] += 1
        if estado["efeitos"]:
            efeito = estado["efeitos"].pop(0)
            if isinstance(efeito, Exception):
                raise efeito
            efeito(driver)

    monkeypatch.setattr(main, "selecionar_paginacao_todas", falsa)
    return estado


def test_garantir_grid_completo_devolve_contagem_quando_tudo_materializa(relogio, paginacao, log):
    driver = DriverFalso(lambda: COMPLETO)

    assert main.garantir_grid_completo(driver, wait=None) == (129, 127)
    assert paginacao["chamadas"] == 1


def test_garantir_grid_completo_repete_paginacao_quando_grid_fica_na_primeira_pagina(relogio, paginacao, log):
    status = {"atual": PRIMEIRA_PAGINA}
    driver = DriverFalso(lambda: status["atual"])

    def segunda_tentativa_funciona(d):
        status["atual"] = COMPLETO

    paginacao["efeitos"] = [lambda d: None, segunda_tentativa_funciona]

    assert main.garantir_grid_completo(driver, wait=None) == (129, 127)
    assert paginacao["chamadas"] == 2


def test_garantir_grid_completo_falha_explicitamente_quando_grid_nao_completa(relogio, paginacao, log):
    driver = DriverFalso(lambda: PRIMEIRA_PAGINA)

    with pytest.raises(main.GridIncompleto) as erro:
        main.garantir_grid_completo(driver, wait=None)

    assert "26" in str(erro.value) and "127" in str(erro.value)
    assert paginacao["chamadas"] == 2


def test_garantir_grid_completo_tolera_erro_na_paginacao_e_tenta_de_novo(relogio, paginacao, log):
    status = {"atual": PRIMEIRA_PAGINA}
    driver = DriverFalso(lambda: status["atual"])

    def agora_funciona(d):
        status["atual"] = COMPLETO

    paginacao["efeitos"] = [main.TimeoutException("select não abriu"), agora_funciona]

    assert main.garantir_grid_completo(driver, wait=None) == (129, 127)
    assert paginacao["chamadas"] == 2


def test_esperar_datagrid_desiste_cedo_se_grid_parado_sem_carregar(relogio, log):
    """Grid em 26/127 sem spinner por 20s = a paginação não pegou. Esperar os
    300s do timeout só atrasa a nova tentativa."""
    driver = DriverFalso(lambda: PRIMEIRA_PAGINA)
    inicio = relogio.time()

    assert main.esperar_datagrid_carregar_todas(driver, timeout=300) == (26, 127)
    assert relogio.time() - inicio < 60


def test_esperar_datagrid_segura_enquanto_grid_ainda_carrega(relogio, log):
    """Com o spinner visível o fetch está em andamento: não é estagnação."""
    inicio = relogio.time()

    def status():
        if relogio.time() - inicio < 90:
            return {"carregadas": 26, "total": 127, "carregando": True}
        return COMPLETO

    driver = DriverFalso(status)

    assert main.esperar_datagrid_carregar_todas(driver, timeout=300) == (129, 127)
    assert relogio.time() - inicio >= 90
