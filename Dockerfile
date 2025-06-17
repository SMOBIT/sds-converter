FROM python:3.11-slim

WORKDIR /app

# System-Abhängigkeiten
RUN apt-get update && apt-get install -y libgl1 libglib2.0-0

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY converter.py .
COPY templates/ templates/

ENTRYPOINT ["python", "converter.py"]
