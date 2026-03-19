# Use uma imagem base com Python e Chrome
FROM python:3.12-slim

# Instalar dependências do sistema
RUN apt-get update && apt-get install -y \
    wget \
    gnupg \
    unzip \
    curl \
    xvfb \
    # Dependências para o Chrome
    libnss3 \
    libnspr4 \
    libatk1.0-0 \
    libatk-bridge2.0-0 \
    libcups2 \
    libdrm2 \
    libdbus-1-3 \
    libatspi2.0-0 \
    libx11-6 \
    libxcomposite1 \
    libxdamage1 \
    libxext6 \
    libxfixes3 \
    libxrandr2 \
    libgbm1 \
    libxcb1 \
    libxkbcommon0 \
    libgtk-3-0 \
    libpango-1.0-0 \
    libcairo2 \
    libasound2 \
    # Limpa cache do apt
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Instala Google Chrome estável
RUN wget -q -O - https://dl-ssl.google.com/linux/linux_signing_key.pub | gpg --dearmor > /etc/apt/trusted.gpg.d/google.gpg \
    && echo "deb [arch=amd64] http://dl.google.com/linux/chrome/deb/ stable main" > /etc/apt/sources.list.d/google-chrome.list \
    && apt-get update \
    && apt-get install -y google-chrome-stable \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Instala ChromeDriver manualmente
RUN CHROME_VERSION=$(google-chrome --version | sed -E "s/.* ([0-9]+)\..*/\1/") \
    && CHROMEDRIVER_VERSION=$(wget -qO- "https://googlechromelabs.github.io/chrome-for-testing/LATEST_RELEASE_$CHROME_VERSION") \
    && wget -q "https://storage.googleapis.com/chrome-for-testing-public/$CHROMEDRIVER_VERSION/linux64/chromedriver-linux64.zip" \
    && unzip chromedriver-linux64.zip \
    && mv chromedriver-linux64/chromedriver /usr/local/bin/ \
    && chmod +x /usr/local/bin/chromedriver \
    && rm -rf chromedriver-linux64.zip chromedriver-linux64

# Criar usuário não-root
RUN useradd -m -s /bin/bash automation

# Define diretório de trabalho
WORKDIR /app

# Copia requirements primeiro
COPY requirements.txt .

# Instala dependências Python com verbose para debug
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt && \
    pip list && \
    python -c "import xlrd; print(f'xlrd {xlrd.__version__} installed successfully')"

# Copia arquivos do projeto
COPY main.py .
COPY docker-entrypoint.sh /
RUN chmod +x /docker-entrypoint.sh

# Criar diretórios necessários
RUN mkdir -p \
    /app/downloads \
    /app/logs \
    /app/relatorios/engenharia \
    /app/relatorios/suprimentos/habitat \
    /app/relatorios/administrativo \
    && chmod -R 777 /app

# Dar permissões ao usuário
RUN chown -R automation:automation /app /home/automation

# Mudar para usuário não-root
USER automation

# Variáveis de ambiente
ENV PYTHONUNBUFFERED=1
ENV DISPLAY=:1
ENV DOWNLOAD_DIR=/app/downloads
ENV LOG_DIR=/app/logs
ENV DOCKER_CONTAINER=true
ENV PATH="/usr/local/bin:${PATH}"

# Define o ponto de entrada e comando
ENTRYPOINT ["/docker-entrypoint.sh"]
CMD ["python3", "main.py"]