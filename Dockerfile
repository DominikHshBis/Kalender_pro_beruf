FROM python:3.11-slim

# Arbeitsverzeichnis im Container
WORKDIR /app

# Systemabhängigkeiten (für google-api-client etc.)
RUN apt-get update && apt-get install -y \
    gcc \
    && rm -rf /var/lib/apt/lists/*

# Requirements zuerst kopieren und installieren
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Restliche Dateien kopieren
COPY . .

# Standard-Startbefehl (wird durch Cron überschrieben)
CMD ["python3", "main.py"]
