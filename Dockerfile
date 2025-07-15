FROM python:3.10

# Instala dependências do sistema
RUN apt-get update && apt-get install -y \
    ghostscript \
    python3-tk \
    python3-opencv \
    libgl1 \
    libglib2.0-0 \
    poppler-utils \
    tesseract-ocr \
    tesseract-ocr-por \
    && rm -rf /var/lib/apt/lists/*

# Cria diretório de trabalho
WORKDIR /app

# Copia dependências Python e instala
COPY requirements.txt .
RUN pip install --upgrade pip
RUN pip install -r requirements.txt

# Copia o restante do código
COPY . .

# Roda o app
CMD ["python", "app.py"]
