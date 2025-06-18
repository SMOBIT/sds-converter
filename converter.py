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
    doc = Document(raw_docx_path)
    sections: dict[str, list] = {}
    current = None
    header_re = re.compile(r"^\s*abschnitt\s*(\d+)\s*[:.-]?", re.I)

    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            m = header_re.match(text)
            if m:
                current = m.group(1)
                sections[current] = []
                continue
            if current:
                sections[current].append(block)
        elif isinstance(block, Table):
            # handle tables similarly
            for row in block.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        m = header_re.match(para.text.strip())
                        if m:
                            current = m.group(1)
                            sections[current] = []
                            break
                    if current and not sections[current]:
                        continue
            if current:
                sections[current].append(block)
    return sections


def merge_into_template(sections, template_path, out_path):
    # Debug: zeigen, welcher Pfad geladen wird
    print(f">>> merge_into_template lädt Template von: {template_path}")
    if not os.path.isfile(template_path):
        print(f"Template not found: {template_path}")
        return
    tpl = Document(template_path)
    body = tpl.element.body
    pattern = re.compile(r"\{+\s*SECTION\s*_?\s*(\d+)\s*\}+", re.I)

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

    # second pass for paragraphs at top level
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

    # process each PDF and then remove it
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

        # PDF entfernen nach erfolgreichem Export
        try:
            os.remove(pdf)
            print(f"Removed input PDF: {pdf}")
        except Exception as e:
            print(f"Could not remove PDF {pdf}: {e}")

        # Roh-DOCX entfernen, falls vorhanden
        if os.path.exists(raw):
            try:
                os.remove(raw)
                print(f"Removed raw DOCX: {raw}")
            except Exception as e:
                print(f"Could not remove raw DOCX {raw}: {e}")

    # abschließende Aufräumaktion: alle verbleibenden Dateien im INPUT_DIR löschen
    for leftover in os.listdir(INPUT_DIR):
        path = os.path.join(INPUT_DIR, leftover)
        if os.path.isfile(path):
            try:
                os.remove(path)
                print(f"Removed leftover file: {path}")
            except Exception as e:
                print(f"Could not remove leftover {path}: {e}")
