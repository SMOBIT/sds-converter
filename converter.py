import os
from pdf2docx import Converter
from docx import Document

# Konstanten für Verzeichnisse im Container
INPUT_DIR = "/app/sample_pdfs"
TEMPLATE_PATH = "/app/templates/master_template.docx"
OUTPUT_DIR = "/app/output"


def pdf_to_raw_docx(pdf_path, raw_docx_path):
    cv = Converter(pdf_path)
    cv.convert(raw_docx_path, start=0, end=None)
    cv.close()


def extract_sections(raw_docx_path):
    doc = Document(raw_docx_path)
    sections = {}
    current = None
    for p in doc.paragraphs:
        text = p.text.strip()
        if text.upper().startswith("ABSCHNITT"):
            sec_num = text.split()[1].rstrip(":")
            current = sec_num
            sections[current] = []
        if current:
            sections[current].append(p)
    return sections


def merge_into_template(sections, template_path, out_path):
    tpl = Document(template_path)
    for sec_num, paras in sections.items():
        placeholder = f"{{{{SECTION_{sec_num}}}}}"
        for p in tpl.paragraphs:
            if placeholder in p.text:
                parent = p._element.getparent()
                idx = parent.index(p._element)
                parent.remove(p._element)
                for para in paras:
                    new_p = tpl.add_paragraph()
                    new_p._element = para._element
                    parent.insert(idx, new_p._element)
                    idx += 1
                break
    tpl.save(out_path)


if __name__ == "__main__":
    # Sicherstellen, dass OUTPUT_DIR existiert
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # Alle PDF-Dateien im INPUT_DIR verarbeiten
    for fname in os.listdir(INPUT_DIR):
        if not fname.lower().endswith('.pdf'):
            continue
        pdf_path = os.path.join(INPUT_DIR, fname)
        raw_docx = os.path.join(OUTPUT_DIR, fname.replace('.pdf', '_raw.docx'))
        final_docx = os.path.join(OUTPUT_DIR, fname.replace('.pdf', '.docx'))

        print(f"Processing {fname}...")
        pdf_to_raw_docx(pdf_path, raw_docx)
        secs = extract_sections(raw_docx)
        merge_into_template(secs, TEMPLATE_PATH, final_docx)
        print(f"Converted {fname} → {os.path.basename(final_docx)}")
