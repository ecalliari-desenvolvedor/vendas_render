# Usa uma imagem oficial leve do Python
FROM python:3.13-slim

# Garante que os logs sejam exibidos imediatamente no console do Render
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1

# Define o diretório de trabalho dentro do container
WORKDIR /app

# Copia o arquivo de dependências para o container
COPY requirements.txt .

# Instala as dependências do Python
RUN pip install --no-cache-dir -r requirements.txt

# Copia o restante do código do projeto para o container
COPY . .

# Garante a existência dos diretórios estáticos
RUN mkdir -p static/tabelas static/vendas static/carteira

# Porta padrão de escuta (o Render substitui pela variável $PORT dinamicamente)
EXPOSE 5000

# Executa com Gunicorn em produção
CMD exec gunicorn --bind 0.0.0.0:${PORT:-5000} --workers 2 --threads 4 app:app
