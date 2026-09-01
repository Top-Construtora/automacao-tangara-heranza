"""Importa o main.py da automação fora do Docker/Windows.

O módulo cria diretórios e fixa o caminho do log na importação; em máquina de
desenvolvimento esses caminhos (/app, C:/Users/...) não existem. Os testes
nunca tocam disco nem Selenium: só exercitam a política de paginação/espera.
"""
import os
import sys
import pathlib

RAIZ = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

try:
    from dotenv import load_dotenv
    load_dotenv(RAIZ / ".env")
except ImportError:
    pass

_makedirs = os.makedirs
os.makedirs = lambda *a, **k: None
try:
    import main  # noqa: E402,F401
finally:
    os.makedirs = _makedirs

import pytest  # noqa: E402


@pytest.fixture
def log(monkeypatch):
    linhas = []
    monkeypatch.setattr(main, "adicionar_ao_log", lambda m, *a, **k: linhas.append(m))
    return linhas
