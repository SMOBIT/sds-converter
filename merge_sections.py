from docxtpl import DocxTemplate
import json

TEMPLATE_PATH = "/data/templates/master_template.docx"
SECTIONS_JSON = "/data/output/sections.json"
OUTPUT_DOCX = "/data/output/merged_output.docx"

def main():
    # Lade die Daten aus der JSON-Datei
    with open(SECTIONS_JSON, encoding="utf-8") as f:
        data = json.load(f)

    # Lade das Word-Template
    doc = DocxTemplate(TEMPLATE_PATH)

    # Erzeuge den Kontext für docxtpl
    context = {
        "SECTION_1": data.get("Abschnitt 1", ""),
        "SECTION_2": data.get("Abschnitt 2", ""),
        "SECTION_3": data.get("Abschnitt 3", ""),
        "SECTION_4": data.get("Abschnitt 4", ""),
        "SECTION_5": data.get("Abschnitt 5", ""),
        "SECTION_6": data.get("Abschnitt 6", ""),
        "SECTION_7": data.get("Abschnitt 7", ""),
        "SECTION_8": data.get("Abschnitt 8", ""),
        "SECTION_9": data.get("Abschnitt 9", ""),
        "SECTION_10": data.get("Abschnitt 10", ""),
        "SECTION_11": data.get("Abschnitt 11", ""),
        "SECTION_12": data.get("Abschnitt 12", ""),
        "SECTION_13": data.get("Abschnitt 13", ""),
        "SECTION_14": data.get("Abschnitt 14", ""),
        "SECTION_15": data.get("Abschnitt 15", ""),
        "SECTION_16": data.get("Abschnitt 16", ""),
    }

    # Ersetze die Platzhalter
    doc.render(context)
    doc.save(OUTPUT_DOCX)

if __name__ == "__main__":
    main()
