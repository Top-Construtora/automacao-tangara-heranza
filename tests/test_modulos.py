"""Seleção de módulos pela variável de ambiente MODULOS (ids separados por
vírgula). Vazia = todos. A Central de Automações usa os mesmos ids."""
import pytest

import main


def test_sem_selecao_roda_todos_na_ordem():
    assert main.selecionar_modulos(main.MODULOS, "") == list(main.MODULOS)
    assert main.selecionar_modulos(main.MODULOS, None) == list(main.MODULOS)


def test_selecao_filtra_mantendo_a_ordem_da_lista():
    ids = [m[0] for m in main.MODULOS]
    escolha = f" {ids[-1]}, {ids[0]} ,"
    assert [m[0] for m in main.selecionar_modulos(main.MODULOS, escolha)] == [ids[0], ids[-1]]


def test_id_desconhecido_aborta_antes_de_abrir_o_navegador():
    with pytest.raises(ValueError, match="nao_existe"):
        main.selecionar_modulos(main.MODULOS, "nao_existe")


def test_ids_sao_estaveis_e_batem_com_a_central():
    assert [m[0] for m in main.MODULOS] == ["cadastro_contratos", "analitico_apropriacoes", "orcado_comprometido", "medido_comprometido", "apropriacoes_insumos", "pedidos_compra", "relacao_solicitacoes", "painel_suprimentos"]
