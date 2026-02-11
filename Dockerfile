# 1. Basis-Image: offizielles Python-Image
FROM python:3.11-slim

# 2. Arbeitsverzeichnis im Container
WORKDIR /app

# 3. Requirements installieren (falls vorhanden)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 4. Restliche Dateien ins Image kopieren
COPY . .

# 5. Standard-Startbefehl (wird von Cron überschrieben)
CMD ["python3", "main.py"]
