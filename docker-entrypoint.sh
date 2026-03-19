#!/bin/bash

# Remove qualquer lock file antigo
rm -f /tmp/.X1-lock

# Inicia o Xvfb no display :1
Xvfb :1 -screen 0 1920x1080x24 -nolisten tcp -nolisten unix &
export DISPLAY=:1

# Aguarda o Xvfb iniciar
sleep 3

# Configurações de ambiente Python
export PYTHONPATH=/app:$PYTHONPATH
export LC_ALL=C.UTF-8
export LANG=C.UTF-8

# Debug: mostra pacotes instalados
echo "=== Pacotes Python instalados ==="
pip list | grep -E "(xlrd|pandas|selenium|openpyxl)"
echo "================================="

# Executa o comando passado (por padrão, python3 main.py)
exec "$@"