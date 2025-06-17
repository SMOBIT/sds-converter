# sds-converter

Diese Open-Source-Pipeline verarbeitet automatisch alle Sicherheitsdatenblätter (PDF) in einem Ordner und mergt deren Inhalte in ein Word-Template.

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

## Installation und Usage

1. **Repository klonen**

   ```bash
   git clone git@github.com:<DeinUser>/sds-converter.git
   cd sds-converter
   ```

2. **Template und PDFs einfügen**

   - Lege Dein `master_template.docx` in das Verzeichnis `templates/`.
   - Kopiere alle PDF-Sicherheitsdatenblätter nach `sample_pdfs/`.

3. **Docker-Image bauen**

   ```bash
   docker build -t sds-converter:latest .
   ```

4. **Output-Ordner anlegen**

   ```bash
   mkdir output
   ```

5. **Converter ausführen**

   ```bash
   docker run --rm \
     -v $(pwd)/sample_pdfs:/app/sample_pdfs \
     -v $(pwd)/templates:/app/templates \
     -v $(pwd)/output:/app/output \
     sds-converter:latest
   ```

   Der Container verarbeitet alle `.pdf`-Dateien in `sample_pdfs/` und legt die fertigen `.docx`-Dateien in `output/` ab.

## Coolify Deployment als Batch-Job

Da der Container nur auf Abruf (per n8n Webhook/Exec-Node) starten und nach der Verarbeitung sofort beenden soll, lege ihn in Coolify als **Batch/One-Shot Job** an, nicht als dauerhaften Web-Service mit Healthchecks.

1. **Im Coolify-Dashboard**: New Project → wähle Dein GitHub-Repo `sds-converter`.
2. **Build Command:**
   ```bash
docker build -t sds-converter:latest .
```
3. **Run Command:**
   ```bash
docker run --rm \
  -v /pfad/host/sample_pdfs:/app/sample_pdfs \
  -v /pfad/host/templates:/app/templates \
  -v /pfad/host/output:/app/output \
  sds-converter:latest
```
4. **Service Type:** Wähle **Job / Command** (One-Shot), damit Coolify kein Healthcheck ausführt und das Container-Exit(0) als Erfolg zählt.
5. **Volumes**: wie oben, für Input, Template und Output.

---

## n8n-Integration
Verwende in n8n einen **Exec-Node** oder Webhook-Trigger, der den Coolify-Job per HTTP API oder direkt ein `docker run` auslöst, wenn neue PDFs in Nextcloud hochgeladen werden. So wird der Batch-Job nur dann gestartet, wenn es nötig ist, und ergibt keine "unhealthy"-Meldungen.
