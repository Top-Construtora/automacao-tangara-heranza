"""Cliente da Edge Function ingest-fluxo-desembolso (sessão HTTP falsa, nada sai da máquina)."""
import pytest
import requests

try:
    from utils.gio_api import ErroApi, GioApi
except ImportError:
    from gio_api import ErroApi, GioApi

URL = "https://exemplo.supabase.co/functions/v1/ingest-fluxo-desembolso"


class Resposta:
    def __init__(self, status, corpo=None, texto=""):
        self.status_code = status
        self._corpo = corpo
        self.text = texto

    def json(self):
        if self._corpo is None:
            raise ValueError("sem json")
        return self._corpo


class SessaoFalsa:
    """Devolve as respostas na ordem; uma exceção na lista é levantada no post."""

    def __init__(self, *respostas):
        self.respostas = list(respostas)
        self.corpos = []
        self.headers = []

    def post(self, url, json=None, headers=None, timeout=None, allow_redirects=True):
        assert allow_redirects is False
        self.corpos.append(json)
        self.headers.append(headers)
        r = self.respostas.pop(0)
        if isinstance(r, Exception):
            raise r
        return r


def api(*respostas):
    sessao = SessaoFalsa(*respostas)
    return GioApi(URL, "chave-secreta", sessao=sessao, pausa=lambda s: None), sessao


def test_exige_url_e_chave():
    with pytest.raises(ErroApi, match="GIO_INGEST_URL e GIO_INGEST_KEY"):
        GioApi("", "x")
    with pytest.raises(ErroApi, match="GIO_INGEST_URL e GIO_INGEST_KEY"):
        GioApi.do_ambiente({"GIO_INGEST_URL": URL})


def test_recusa_http_fora_de_localhost_e_aceita_local():
    with pytest.raises(ErroApi, match="https"):
        GioApi("http://exemplo.com/f", "x")
    GioApi("http://127.0.0.1:54321/functions/v1/ingest-fluxo-desembolso", "x")
    GioApi("http://localhost:54321/functions/v1/ingest-fluxo-desembolso", "x")


def test_chave_vai_no_header_e_nunca_na_mensagem_de_erro():
    cliente, sessao = api(Resposta(400, {"error": "Sienge inválido"}))
    with pytest.raises(ErroApi) as erro:
        cliente.iniciar("top", "a.xlsx", [["x"]])
    assert sessao.headers[0] == {"x-api-key": "chave-secreta"}
    assert "chave-secreta" not in str(erro.value)


def test_conferir_acesso_ok_e_401_explicado():
    cliente, sessao = api(Resposta(200, {"ok": True}))
    cliente.conferir_acesso()
    assert sessao.corpos == [{"acao": "verificar"}]
    cliente, _ = api(Resposta(401, {"error": "x-api-key inválida"}))
    with pytest.raises(ErroApi, match="GIO_INGEST_KEY") as erro:
        cliente.conferir_acesso()
    assert erro.value.status == 401


def test_iniciar_manda_corpo_e_valida_resposta():
    cliente, sessao = api(Resposta(200, {"envio_id": "e-1", "inicio_dados": 3}))
    assert cliente.iniciar("habitat", "V.xlsx", [["a"], ["b"]]) == {"envio_id": "e-1", "inicio_dados": 3}
    assert sessao.corpos == [{"acao": "iniciar", "sienge": "habitat", "arquivo_nome": "V.xlsx",
                              "primeiras_linhas": [["a"], ["b"]]}]


@pytest.mark.parametrize("corpo", [{"inicio_dados": 3}, {"envio_id": "", "inicio_dados": 3},
                                   {"envio_id": "e", "inicio_dados": "3"}, {"envio_id": "e", "inicio_dados": -1},
                                   {"envio_id": "e", "inicio_dados": True}])
def test_iniciar_recusa_resposta_incompleta(corpo):
    cliente, _ = api(Resposta(200, corpo))
    with pytest.raises(ErroApi, match="iniciar"):
        cliente.iniciar("top", "a.xlsx", [])


def test_lote_manda_seq_e_linhas():
    cliente, sessao = api(Resposta(200, {"seq": 2, "lancamentos": 10, "ignoradas": 1}))
    assert cliente.lote("e-1", 2, [[1, "a"]])["lancamentos"] == 10
    assert sessao.corpos == [{"acao": "lote", "envio_id": "e-1", "seq": 2, "linhas": [[1, "a"]]}]


def test_4xx_nao_tenta_de_novo():
    cliente, sessao = api(Resposta(400, {"error": "seq inválido"}), Resposta(200, {}))
    with pytest.raises(ErroApi, match="seq inválido") as erro:
        cliente.lote("e-1", 1, [])
    assert erro.value.status == 400 and len(sessao.corpos) == 1


def test_500_e_rede_tentam_de_novo():
    cliente, sessao = api(Resposta(500, {"error": "x"}), Resposta(200, {"seq": 1}))
    assert cliente.lote("e-1", 1, []) == {"seq": 1}
    assert len(sessao.corpos) == 2
    cliente, sessao = api(requests.Timeout(), Resposta(200, {"seq": 1}))
    assert cliente.lote("e-1", 1, []) == {"seq": 1}
    cliente, _ = api(Resposta(500, {"error": "x"}), Resposta(503, None, "fora"))
    with pytest.raises(ErroApi, match="503") as erro:
        cliente.lote("e-1", 1, [])
    assert erro.value.status == 503


def test_redirecionamento_recusado():
    cliente, sessao = api(Resposta(302, None, ""), Resposta(200, {}))
    with pytest.raises(ErroApi, match="redirecionamento"):
        cliente.lote("e-1", 1, [])
    assert len(sessao.corpos) == 1


def test_200_sem_json_e_erro():
    cliente, _ = api(Resposta(200, None, "<html>"))
    with pytest.raises(ErroApi, match="JSON"):
        cliente.lote("e-1", 1, [])


def test_concluir_ok():
    cliente, sessao = api(Resposta(200, {"obras_reconhecidas": ["A"]}))
    assert cliente.concluir("e-1", 3) == {"obras_reconhecidas": ["A"]}
    assert sessao.corpos == [{"acao": "concluir", "envio_id": "e-1", "total_lotes": 3}]


def test_concluir_500_depois_envio_nao_encontrado_e_sucesso():
    nao_achou = {"error": "Envio não encontrado (expirado ou já concluído)"}
    cliente, _ = api(Resposta(500, {"error": "x"}), Resposta(400, nao_achou))
    assert cliente.concluir("e-1", 3) == {"concluido_por_reenvio": True}
    cliente, _ = api(requests.ConnectionError(), Resposta(400, nao_achou))
    assert cliente.concluir("e-1", 3) == {"concluido_por_reenvio": True}


def test_concluir_envio_nao_encontrado_de_primeira_e_erro():
    cliente, _ = api(Resposta(400, {"error": "Envio não encontrado (expirado ou já concluído)"}))
    with pytest.raises(ErroApi, match="Envio não encontrado"):
        cliente.concluir("e-1", 3)


def test_lote_nao_herda_a_regra_do_concluir():
    cliente, _ = api(Resposta(500, {"error": "x"}), Resposta(400, {"error": "Envio não encontrado"}))
    with pytest.raises(ErroApi, match="Envio não encontrado"):
        cliente.lote("e-1", 1, [])


def test_conferir_acesso_404_explicado():
    cliente, _ = api(Resposta(404, {"error": "Not found"}))
    with pytest.raises(ErroApi, match="404") as erro:
        cliente.conferir_acesso()
    assert erro.value.status == 404


def test_conferir_acesso_200_sem_ok_e_erro():
    cliente, _ = api(Resposta(200, {"ok": False}))
    with pytest.raises(ErroApi, match="verificar"):
        cliente.conferir_acesso()


def test_concluir_500_duas_vezes_e_erro_sem_reenvio():
    cliente, sessao = api(Resposta(500, {"error": "x"}), Resposta(500, {"error": "y"}))
    with pytest.raises(ErroApi) as erro:
        cliente.concluir("e-1", 3)
    assert erro.value.status == 500 and len(sessao.corpos) == 2
