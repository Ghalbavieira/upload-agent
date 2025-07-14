FROM python:3.11-slim

# Instala dependências do sistema
RUN apt-get update && \
    apt-get install -y ghostscript python3-tk poppler-utils && \
    rm -rf /var/lib/apt/lists/*

# Cria diretório de trabalho
WORKDIR /app

# Copia os arquivos do projeto
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py .

# Comando para rodar
CMD ["python", "app.py"]
