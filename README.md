# sds-converter

Dieses Repository enthält eine Docker-basierte Pipeline, um Sicherheitsdatenblätter (PDF) in DOCX umzuwandeln und in ein Master-Template zu mergen.

## Installation

```bash
git clone git@github.com:<DeinUser>/sds-converter.git
cd sds-converter
``` 

## Usage

1. Baue das Docker-Image:
   ```bash
docker build -t sds-converter:latest .
```
2. Erstelle die Verzeichnisse für Input, Templates und Output:
   ```bash
mkdir templates sample_pdfs output
# Lege master_template.docx in templates/
# Lege PDFs in sample_pdfs/
```
3. Starte den Converter:
   ```bash
docker run --rm \
  -v $(pwd)/sample_pdfs:/app/sample_pdfs \
  -v $(pwd)/templates:/app/templates \
  -v $(pwd)/output:/app/output \
  sds-converter:latest \
  sample_pdfs templates output
```

## n8n Integration

- Verwende einen Exec-Node, um das oben genannte Docker-Command auszuführen.
