# Docker Setup para Automação TANGARA

Este projeto containeriza a automação do SIENGE TANGARA usando Docker, permitindo execução em qualquer ambiente sem necessidade de instalação local de dependências.

## 📋 Pré-requisitos

- Docker instalado (versão 20.10 ou superior)
- Docker Compose instalado (versão 1.29 ou superior)

## 🚀 Configuração Rápida

### 1. Clone o projeto e navegue até o diretório

```bash
cd /caminho/para/o/projeto
```

### 2. Configure as credenciais

Crie um arquivo `.env` baseado no exemplo:

```bash
cp .env.example .env
```

Edite o arquivo `.env` com suas credenciais:

```env
TANGARA_USERNAME=seu_usuario
TANGARA_PASSWORD=sua_senha
TANGARA_EMAIL=seu_email@empresa.com
TANGARA_EMAIL_PASSWORD=sua_senha_email
```

### 3. Estrutura de diretórios

O Docker criará automaticamente os diretórios necessários:

```
projeto/
├── docker-compose.yml
├── Dockerfile
├── .dockerignore
├── .env
├── main.py
├── main_docker.py     # Versão adaptada para Docker
├── requirements.txt
├── downloads/         # Arquivos baixados aparecerão aqui
├── relatorios/        # Relatórios gerados
│   ├── ENGENHARIA/
│   ├── SUPRIMENTOS/
│   │   └── TANGARA/
│   └── ADMINISTRATIVO/
└── logs/              # Logs de execução
```

## 🔧 Uso

### Construir a imagem Docker

```bash
docker-compose build
```

### Executar a automação

```bash
docker-compose up
```

### Executar em segundo plano

```bash
docker-compose up -d
```

### Ver logs em tempo real

```bash
docker-compose logs -f
```

### Parar a execução

```bash
docker-compose down
```

## 📁 Arquivos Gerados

Os arquivos baixados e relatórios gerados estarão disponíveis nos diretórios mapeados:

- **Downloads**: `./downloads/`
- **Relatórios**: `./relatorios/`
- **Logs**: `./logs/`

## 🐛 Debugging

### Acessar o container em modo interativo

Descomente as linhas de debug no `docker-compose.yml`:

```yaml
stdin_open: true
tty: true
command: /bin/bash
```

Então execute:

```bash
docker-compose run --rm tangara-automation /bin/bash
```

### Executar com logs detalhados

```bash
docker-compose up --no-log-prefix
```

## ⚠️ Considerações Importantes

1. **Conversão XLS para XLSX**: Como o `pywin32` não funciona em Linux, implementei uma conversão alternativa usando `pandas`. Caso encontre problemas, os arquivos XLS serão mantidos.

2. **Modo Headless**: O Chrome roda em modo headless (sem interface gráfica) no Docker. Isso é necessário para containers.

3. **Captcha**: Se o sistema solicitar captcha, será necessário implementar soluções alternativas ou executar localmente.

4. **Performance**: A execução em container pode ser ligeiramente mais lenta que a execução local.

## 🔐 Segurança

- **Nunca** commite o arquivo `.env` com credenciais reais
- Use secrets do Docker Swarm ou Kubernetes em produção
- Considere usar um gerenciador de senhas/secrets

## 📝 Modificações do Código Original

O arquivo `main_docker.py` contém as seguintes adaptações:

1. Remoção de dependências Windows-specific (`pywin32`)
2. Adaptação de caminhos para Linux
3. Chrome configurado para modo headless
4. Conversão XLS usando `pandas` ao invés de COM
5. Logs mais verbosos para debugging

## 🆘 Troubleshooting

### Erro: "Chrome failed to start"

```bash
# Aumentar memória compartilhada
docker-compose down
docker-compose up -d
```

### Erro: "Permission denied"

```bash
# Dar permissões aos diretórios
chmod -R 755 downloads/ relatorios/ logs/
```

### Container para imediatamente

Verifique os logs:

```bash
docker-compose logs --tail=50
```

## 🚀 Melhorias Futuras

- [ ] Implementar agendamento com cron
- [ ] Adicionar notificações (email/slack)
- [ ] Implementar retry automático em caso de falha
- [ ] Adicionar suporte para múltiplas obras simultâneas