FROM python:3.11-slim

WORKDIR /app

# System-Abhängigkeiten
RUN apt-get update && apt-get install -y libgl1 libglib2.0-0

# Python-Abhängigkeiten
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

# App-Code
COPY converter.py merge_sections.py ./
COPY templates/ ./templates/
COPY sample_pdfs/ ./sample_pdfs/

# Output-Verzeichnis
RUN mkdir -p /app/output

# ENTRYPOINT bleibt "python", CMD ist default auf converter.py
ENTRYPOINT ["python"]
CMD ["converter.py"]
