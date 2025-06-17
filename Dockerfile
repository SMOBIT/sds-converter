FROM python:3.11-slim

WORKDIR /app

# System-Abhängigkeiten installieren
RUN apt-get update && apt-get install -y libgl1 libglib2.0-0

# Python-Abhängigkeiten installieren
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

# App-Code und Verzeichnisse kopieren
COPY converter.py ./
COPY templates/ ./templates/
COPY sample_pdfs/ ./sample_pdfs/

# Sicherstellen, dass das Output-Verzeichnis existiert
RUN mkdir -p /app/output

# Beim Start alle PDFs im Ordner verarbeiten
ENTRYPOINT ["python", "converter.py"]
