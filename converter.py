import os
from pdf2docx import Converter
from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.shared import Inches
from PIL import Image

# Container-Verzeichnisse
INPUT_DIR = "/app/sample_pdfs"
TEMPLATE_PATH = "/app/templates/master_template.docx"
OUTPUT_DIR = "/app/output"
ICONS_DIR = "/app/icons"

# Liste unterstützter Icon-Erweiterungen
ICON_EXTENSIONS = ['png', 'jpg', 'jpeg']


def get_image_size_inches(path):
    """
    Berechnet die Bildgröße in Inches basierend auf Pixelgröße und DPI.
    """
    img = Image.open(path)
    dpi = img.info.get('dpi', (96, 96))
    width_px, height_px = img.size
    return width_px / dpi[0], height_px / dpi[1]


def pdf_to_raw_docx(pdf_path, raw_docx_path):
    """
    Konvertiert eine PDF-Datei zu einem Roh-DOCX und extrahiert alle Rasterbilder.
    """
    img_folder = raw_docx_path.replace('.docx', '_images')
    os.makedirs(img_folder, exist_ok=True)

    cv = Converter(pdf_path)
    cv.convert(
        raw_docx_path,
        start=0,
        end=None,
        image_folder=img_folder,
        extract_images=True
    )
    cv.close()


def iter_block_items(doc):
    """
    Liefert Absätze und Tabellen in Dokumentreihenfolge.
    """
    for child in doc.element.body:
        if isinstance(child, CT_P):
            yield Paragraph(child, doc)
        elif isinstance(child, CT_Tbl):
            yield Table(child, doc)
        else:
            yield child


def extract_sections(raw_docx_path):
    """
    Liest Abschnitte basierend auf Überschrift 'ABSCHNITT <Nummer>' aus dem Roh-DOCX.
    """
    doc = Document(raw_docx_path)
    sections = {}
    current = None
    for block in iter_block_items(doc):
        text = block.text.strip() if isinstance(block, Paragraph) else ''
        if text.upper().startswith("ABSCHNITT"):
            sec_num = text.split()[1].rstrip(":")
            current = sec_num
            sections[current] = []
        if current:
            sections[current].append(block)
    return sections


def merge_into_template(sections, template_path, out_path):
    """
    Fügt extrahierte Sektionen ins Master-Template ein; ergänzt bei Bedarf Fallback-Icons.
    Unterstützt PNG und JPG.
    """
    if not os.path.isfile(template_path):
        print(f"WARNING: Template not found at {template_path}, skipping merge.")
        return
    tpl = Document(template_path)
    body = tpl.element.body

    for sec_num, blocks in sections.items():
        placeholder = f"{{SECTION_{sec_num}}}"
        for para in tpl.paragraphs:
            if placeholder in para.text:
                idx = body.index(para._element)
                body.remove(para._element)
                # Roh-Inhalte einsetzen
                for block in blocks:
                    element = getattr(block, '_element', block)
                    body.insert(idx, element)
                    idx += 1
                # Fallback-Icon suchen und einfügen
                icon_path = None
                for ext in ICON_EXTENSIONS:
                    candidate = os.path.join(ICONS_DIR, f"GHS{sec_num}.{ext}")
                    if os.path.isfile(candidate):
                        icon_path = candidate
                        break
                if icon_path:
                    width_in, _ = get_image_size_inches(icon_path)
                    pic_par = tpl.add_paragraph()
                    run = pic_par.add_run()
                    run.add_picture(icon_path, width=Inches(width_in))
                    body.insert(idx, pic_par._p)
                break
    tpl.save(out_path)

if __name__ == "__main__":
    # Verzeichnisse sicherstellen
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(ICONS_DIR, exist_ok=True)

    # Verarbeitung aller PDFs
    for fname in os.listdir(INPUT_DIR):
        if not fname.lower().endswith('.pdf'):
            continue
        pdf_path = os.path.join(INPUT_DIR, fname)
        raw_docx = os.path.join(OUTPUT_DIR, fname.replace('.pdf', '_raw.docx'))
        final_docx = os.path.join(OUTPUT_DIR, fname.replace('.pdf', '.docx'))

        print(f"Processing {fname}...")
        pdf_to_raw_docx(pdf_path, raw_docx)
        img_folder = raw_docx.replace('.docx', '_images')
        print(f"Images folder: {img_folder} -> {'found' if os.path.isdir(img_folder) else 'missing'}")
        sections = extract_sections(raw_docx)
        merge_into_template(sections, TEMPLATE_PATH, final_docx)
        print(f"Converted {fname} → {os.path.basename(final_docx)}")
