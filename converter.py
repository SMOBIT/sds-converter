import os
from pdf2docx import Converter
from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph
from copy import deepcopy
import re
from docx.shared import Inches
from PIL import Image

# Standardpfade (werden im Container verwendet)
DEFAULT_INPUT_DIR = "/data/sample_pdfs"
DEFAULT_TEMPLATE_PATH = "/data/templates/master_template.docx"
DEFAULT_OUTPUT_DIR = "/data/output"
DEFAULT_ICONS_DIR = "/data/icons"

# Pfade ggf. überschreiben, damit das Skript auch lokal lauffähig ist
INPUT_DIR = os.environ.get("INPUT_DIR", DEFAULT_INPUT_DIR)
TEMPLATE_PATH = os.environ.get("TEMPLATE_PATH", DEFAULT_TEMPLATE_PATH)
OUTPUT_DIR = os.environ.get("OUTPUT_DIR", DEFAULT_OUTPUT_DIR)
ICONS_DIR = os.environ.get("ICONS_DIR", DEFAULT_ICONS_DIR)

# Debug-Ausgabe des verwendeten Template-Pfads
print(f">>> Verwende TEMPLATE_PATH: {TEMPLATE_PATH}")


def get_image_size_inches(path):
    img = Image.open(path)
    dpi = img.info.get('dpi', (96, 96))
    w_px, h_px = img.size
    return w_px / dpi[0], h_px / dpi[1]


def pdf_to_raw_docx(pdf_path, raw_docx_path):
    img_folder = raw_docx_path.replace('.docx', '_images')
    os.makedirs(img_folder, exist_ok=True)
    cv = Converter(pdf_path)
    cv.convert(raw_docx_path,
               start=0, end=None,
               image_folder=img_folder,
               extract_images=True)
    cv.close()


def iter_block_items(parent):
    # yield paragraphs and tables only
    for child in parent.element.body:
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def extract_sections(raw_docx_path):
    """
    Teilt ein rohes DOCX in Abschnitte auf, die durch 'Abschnitt X:' in
    Absätzen oder Tabellenüberschriften markiert werden.
    """
    doc = Document(raw_docx_path)
    sections: dict[str, list] = {}
    current = None
    header_re = re.compile(r"^\s*abschnitt\s*(\d+)\s*[:.-]?", re.I)

    for block in iter_block_items(doc):
        # Paragraph-Überschrift
        if isinstance(block, Paragraph):
            text = block.text.strip()
            m = header_re.match(text)
            if m:
                current = m.group(1)
                sections[current] = []
                continue  # Überschrift nicht als Inhalt
            if current:
                sections[current].append(block)

        # Tabellenblock
        elif isinstance(block, Table):
            # Suche nach einer Abschnitts-Überschrift innerhalb der Tabelle
            found_header = False
            header_row_index = None
            header_num = None
            for ri, row in enumerate(block.rows):
                for cell in row.cells:
                    for para in cell.paragraphs:
                        m = header_re.match(para.text.strip())
                        if m:
                            header_num = m.group(1)
                            found_header = True
                            header_row_index = ri
                            break
                    if found_header:
                        break
                if found_header:
                    break

            if found_header:
                current = header_num
                sections[current] = []

                # kopiere Tabelle ohne Überschriftszeilen
                tbl_elem = deepcopy(block._element)
                for _ in range(header_row_index + 1):
                    tbl_elem.remove(tbl_elem.tr_lst[0])
                if len(tbl_elem.tr_lst) > 0:
                    sections[current].append(Table(tbl_elem, doc))
            else:
                # Ansonsten gesamte Tabelle dem aktuellen Abschnitt hinzufügen
                if current:
                    sections[current].append(block)

    return sections


def merge_into_template(sections, template_path, out_path):
    print(f">>> merge_into_template lädt Template von: {template_path}")
    if not os.path.isfile(template_path):
        print(f"Template not found: {template_path}")
        return
    tpl = Document(template_path)
    body = tpl.element.body
    pattern = re.compile(r"\{+\s*SECTION\s*_?\s*(\d+)\s*\}+", re.I)

    # Erster Pass: Einfügen direkt in Platzhalter-Absätzen
    for p in tpl.element.xpath('.//w:p'):
        texts = [t.text or '' for t in p.xpath('.//w:t')]
        text = ''.join(texts)
        m = pattern.search(text)
        if not m:
            continue
        num = m.group(1)
        parent = p.getparent()
        idx = parent.index(p)
        parent.remove(p)
        for b in sections.get(num, []):
            elem = getattr(b, '_element', b)
            parent.insert(idx, deepcopy(elem))
            idx += 1

    # Zweiter Pass: oberste Ebene
    pattern2 = re.compile(r"{SECTION_(\d+)}")
    for block in list(iter_block_items(tpl)):
        if not isinstance(block, Paragraph):
            continue
        text = block.text
        m2 = pattern2.search(text)
        if not m2:
            continue
        num = m2.group(1)
        idx = body.index(block._element)
        body.remove(block._element)
        for b in sections.get(num, []):
            elem = getattr(b, '_element', b)
            body.insert(idx, deepcopy(elem))
            idx += 1

    tpl.save(out_path)


if __name__ == '__main__':
    os.makedirs(INPUT_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(ICONS_DIR, exist_ok=True)

    for f in os.listdir(INPUT_DIR):
        if not f.lower().endswith('.pdf'):
            continue
        pdf = os.path.join(INPUT_DIR, f)
        base, _ = os.path.splitext(f)
        raw = os.path.join(OUTPUT_DIR, f"{base}_raw.docx")
        final = os.path.join(OUTPUT_DIR, f"{base}.docx")

        print('Processing', f)
        pdf_to_raw_docx(pdf, raw)
        secs = extract_sections(raw)
        merge_into_template(secs, TEMPLATE_PATH, final)
        print('Saved', final)

        # Dateien löschen: nur PDF & Raw-DOCX
                # Nur Eingangs-PDF löschen, keine Dateien aus OUTPUT_DIR
        if os.path.exists(pdf):
            try:
                os.remove(pdf)
                print(f"Removed input PDF: {pdf}")
            except Exception as e:
                print(f"Could not remove PDF {pdf}: {e}")
