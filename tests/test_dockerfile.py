"""O Dockerfile copia os arquivos um a um: módulo importado pelo main.py que ficar fora
da imagem só aparece em produção ("No module named ..."), nunca no pytest local."""
import pathlib
import re

RAIZ = pathlib.Path(__file__).resolve().parents[1]


def test_dockerfile_copia_os_modulos_do_envio_ao_gio():
    copiados = set()
    for linha in (RAIZ / "Dockerfile").read_text(encoding="utf-8").splitlines():
        m = re.match(r"\s*COPY\s+(.+?)\s+\S+\s*$", linha)
        if m:
            copiados.update(m.group(1).split())
    for arquivo in ("main.py", "gio_api.py", "fluxo_desembolso.py", "painel_suprimentos_envio.py"):
        assert arquivo in copiados, f"{arquivo} fora do COPY do Dockerfile"
