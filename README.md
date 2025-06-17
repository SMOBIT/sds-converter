# sds-converter

Dieses Repository enthält eine Docker-basierte Pipeline, um Sicherheitsdatenblätter (PDF) in DOCX umzuwandeln und in ein Master-Template zu mergen.

## Verzeichnisstruktur nach dem Klonen

```bash
sds-converter/
├── Dockerfile
├── requirements.txt
├── converter.py
├── README.md
├── templates/
│   └── .gitkeep
└── sample_pdfs/
    └── .gitkeep
```

## Installation

```bash
git clone git@github.com:<DeinUser>/sds-converter.git
cd sds-converter
``` 

## Usage

1. Füge Dein `master_template.docx` in das Verzeichnis `templates/` ein.
2. Lege Deine PDF-Dateien in `sample_pdfs/`.
3. Baue das Docker-Image:
   ```bash
docker build -t sds-converter:latest .
```
4. Erstelle ein Verzeichnis für das Ausgabeziel:
   ```bash
mkdir output
```
5. Starte den Converter:
   ```bash
docker run --rm \
  -v $(pwd)/sample_pdfs:/app/sample_pdfs \
  -v $(pwd)/templates:/app/templates \
  -v $(pwd)/output:/app/output \
  sds-converter:latest \
  sample_pdfs templates output
```

## n8n Integration

- Verwende einen Exec-Node, um das oben genannte Docker-Command auszuführen, z.B.:
  ```bash
  docker run --rm \
    -v /pfad/host/sample_pdfs:/app/sample_pdfs \
    -v /pfad/host/templates:/app/templates \
    -v /pfad/host/output:/app/output \
    sds-converter:latest \
    sample_pdfs templates output
  ```
