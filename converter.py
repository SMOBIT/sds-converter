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


def get_image_size_inches(path):
    img = Image.open(path)
    dpi = img.info.get('dpi', (96, 96))
    width_px, height_px = img.size
    return width_px / dpi[0], height_px / dpi[1]


def pdf_to_raw_docx(pdf_path, raw_docx_path):
    # Lege einen temporären Image-Ordner an
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


def iter_block_items(parent):
    """
    Yield each paragraph, table, or other block item in document order.
    """
    for child in parent.element.body:
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)
        else:
            yield child


def extract_sections(raw_docx_path):
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
    if not os.path.isfile(template_path):
        print(f"WARNING: Template not found at {template_path}, skipping merge.")
        return
    tpl = Document(template_path)
    body = tpl.element.body
    for sec_num, blocks in sections.items():
        placeholder = f"{{{{SECTION_{sec_num}}}}}"
        for para in tpl.paragraphs:
            if placeholder in para.text:
                idx = body.index(para._element)
                body.remove(para._element)
                # Inhalte aus Roh-DOCX einfügen
                for block in blocks:
                    element = getattr(block, '_element', block)
                    body.insert(idx, element)
                    idx += 1
                # Fallback-Icon, wenn keine Grafik gefunden wurde
                icon_filename = f"GHS{sec_num}.png"
                icon_path = os.path.join(ICONS_DIR, icon_filename)
                if os.path.isfile(icon_path):
                    if not any(hasattr(b, 'inline_shapes') and b.inline_shapes for b in blocks):
                        w_in, _ = get_image_size_inches(icon_path)
                        pic_par = tpl.add_paragraph()
                        run = pic_par.add_run()
                        run.add_picture(icon_path, width=Inches(w_in))
                        body.insert(idx, pic_par._p)
                        idx += 1
                break
    tpl.save(out_path)

if __name__ == "__main__":
    # Ausgabeordner anlegen
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    # Icons-Verzeichnis (für Fallback-Bilder) prüfen
    os.makedirs(ICONS_DIR, exist_ok=True)

    # Alle PDF-Dateien im Input-Ordner verarbeiten
    for fname in os.listdir(INPUT_DIR):
        if not fname.lower().endswith('.pdf'):
            continue
        pdf_path = os.path.join(INPUT_DIR, fname)
        raw_docx = os.path.join(OUTPUT_DIR, fname.replace('.pdf', '_raw.docx'))
        final_docx = os.path.join(OUTPUT_DIR, fname.replace('.pdf', '.docx'))

        print(f"Processing {fname}...")
        pdf_to_raw_docx(pdf_path, raw_docx)
        # Nach Konvertierung prüfen, ob Bilder-Ordner existiert
        img_folder = raw_docx.replace('.docx', '_images')
        if os.path.isdir(img_folder):
            print(f"Images extracted to {img_folder}")
        else:
            print(f"WARNING: No images folder created for {fname}")
        # Merge nur durchführen, wenn Template vorhanden
        secs = extract_sections(raw_docx)
        merge_into_template(secs, TEMPLATE_PATH, final_docx)
        print(f"Converted {fname} → {os.path.basename(final_docx)}") == "__main__":
    # Ausgabeordner anlegen
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(ICONS_DIR, exist_ok=True)

    # Alle PDF-Dateien im Input-Ordner verarbeiten
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
```python
import os
from pdf2docx import Converter
from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.shared import Inches

# Container-Verzeichnisse
INPUT_DIR = "/app/sample_pdfs"
TEMPLATE_PATH = "/app/templates/master_template.docx"
OUTPUT_DIR = "/app/output"


def pdf_to_raw_docx(pdf_path, raw_docx_path):
    # Lege einen temporären Image-Ordner an
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


def iter_block_items(parent):
    """
    Yield each paragraph, table, or other block item in document order.
    """
    for child in parent.element.body:
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)
        else:
            yield child


def extract_sections(raw_docx_path):
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
    tpl = Document(template_path)
    body = tpl.element.body
    for sec_num, blocks in sections.items():
        placeholder = f"{{SECTION_{sec_num}}}"
        for para in tpl.paragraphs:
            if placeholder in para.text:
                idx = body.index(para._element)
                body.remove(para._element)
                # Inhalte aus Roh-DOCX einfügen
                for block in blocks:
                    element = getattr(block, '_element', block)
                    body.insert(idx, element)
                    idx += 1
                # Fallback-Icon, wenn keine Bilder gefunden wurden
                icon_path = f"icons/GHS{sec_num}.png"
                if not any(isinstance(b, Paragraph) and f"GHS{sec_num}" in b.text for b in blocks):
                    pic_par = tpl.add_paragraph()
                    run = pic_par.add_run()
                    run.add_picture(icon_path, width=Inches(0.5))
                    body.insert(idx, pic_par._p)
                    idx += 1
                break
    tpl.save(out_path)


if __name__ == "__main__":
    # Ausgabeordner anlegen
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # Alle PDF-Dateien im Input-Ordner verarbeiten
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
```python
import os
from pdf2docx import Converter
from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph

# Container-Verzeichnisse
INPUT_DIR = "/app/sample_pdfs"
TEMPLATE_PATH = "/app/templates/master_template.docx"
OUTPUT_DIR = "/app/output"


def pdf_to_raw_docx(pdf_path, raw_docx_path):
    # Lege einen temporären Image-Ordner an
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


def iter_block_items(parent):
    """
    Yield each paragraph, table, or other block item in document order.
    """
    for child in parent.element.body:
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)
        else:
            yield child


def extract_sections(raw_docx_path):
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
    tpl = Document(template_path)
    body = tpl.element.body
    for sec_num, blocks in sections.items():
        placeholder = f"{{{{SECTION_{sec_num}}}}}"
        for para in tpl.paragraphs:
            if placeholder in para.text:
                idx = body.index(para._element)
                body.remove(para._element)
                for block in blocks:
                    element = getattr(block, '_element', block)
                    body.insert(idx, element)
                    idx += 1
                break
    tpl.save(out_path)


if __name__ == "__main__":
    # Ausgabeordner anlegen
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # Alle PDF-Dateien im Input-Ordner verarbeiten
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
``` "__main__":
    # Ausgabeordner anlegen
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # Alle PDF-Dateien im Input-Ordner verarbeiten
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
```python
import os
from pdf2docx import Converter
from docx import Document

# Container-Verzeichnisse
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
    # Ausgabeordner anlegen
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    # Alle PDF-Dateien im Input-Ordner verarbeiten
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
