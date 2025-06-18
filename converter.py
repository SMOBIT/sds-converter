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

    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            m = header_re.match(text)
            if m:
                current = m.group(1)
                sections[current] = []
                continue  # skip heading itself
            if current:
                sections[current].append(block)
        elif isinstance(block, Table):
            found_idx = None
            for i, row in enumerate(block.rows):
                for cell in row.cells:
                    for para in cell.paragraphs:
                        m = header_re.match(para.text.strip())
                        if m:
                            current = m.group(1)
                            sections[current] = []
                            found_idx = i
                            break
                    if found_idx is not None:
                        break
                if found_idx is not None:
                    break
            if found_idx is not None:
                tbl_elem = deepcopy(block._element)
                rows = tbl_elem.findall('./w:tr', ns)
                for r in rows[:found_idx+1]:
                    tbl_elem.remove(r)
                if tbl_elem.findall('./w:tr', ns):
                    sections[current].append(tbl_elem)
            elif current:
                sections[current].append(block)
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            m = header_re.match(text)
            if m:
                current = m.group(1)
                sections[current] = []
                continue  # skip heading itself
            if current:
                sections[current].append(block)
        elif isinstance(block, Table):
            found_idx = None
            for i, row in enumerate(block.rows):
                for cell in row.cells:
                    for para in cell.paragraphs:
                        m = header_re.match(para.text.strip())
                        if m:
                            current = m.group(1)
                            sections[current] = []
                            found_idx = i
                            break
                    if found_idx is not None:
                        break
                if found_idx is not None:
                    break
            if found_idx is not None:
                tbl_elem = deepcopy(block._element)
                rows = tbl_elem.findall('./w:tr', ns)
                for r in rows[:found_idx+1]:
                    tbl_elem.remove(r)
                if tbl_elem.findall('./w:tr', ns):
                    sections[current].append(tbl_elem)
            elif current:
                sections[current].append(block)
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            m = header_re.match(text)
            if m:
                current = m.group(1)
                sections[current] = []
            if current:
                sections[current].append(block)
        elif isinstance(block, Table):
            found = False
            for row in block.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        m = header_re.match(para.text.strip())
                        if m:
                            current = m.group(1)
                            sections[current] = []
                            found = True
                            break
                    if found:
                        break
                if found:
                    break
            if current:
                sections[current].append(block)
    header_re = re.compile(r"^\s*abschnitt\s*(\d+)", re.I)
    for block in iter_block_items(doc):
        text = block.text.strip() if isinstance(block, Paragraph) else ''
        m = header_re.match(text)
        if m:
            current = m.group(1)
            sections[current] = []
        if current:
            sections[current].append(block)
    return sections


def merge_into_template(sections, template_path, out_path):
    if not os.path.isfile(template_path):
        print(f"Template not found: {template_path}")
        return
    tpl = Document(template_path)
    body = tpl.element.body
    # Regex zum Finden von Platzhaltern. Erlaubt sind ein oder mehrere
    # geschweifte Klammern sowie optionale Leerzeichen innerhalb des
    # Platzhalters, z.B. "{SECTION_1}", "{{ SECTION 1 }}" usw.
    pattern = re.compile(r"\{+\s*SECTION\s*_?\s*(\d+)\s*\}+", re.I)

    # Regex zum Finden von Platzhaltern wie {SECTION_1} oder {{SECTION_1}}
    # (manche Templates nutzen doppelte geschweifte Klammern)
    pattern = re.compile(r"\{{1,2}SECTION_(\d+)\}{1,2}")


    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

    # iterate over all paragraph elements, auch innerhalb von Tabellen
    for p in tpl.element.xpath('.//w:p', namespaces=ns):
        texts = [t.text or '' for t in p.xpath('.//w:t', namespaces=ns)]
        text = ''.join(texts)
        m = pattern.search(text)
        if not m:

            continue
        num = m.group(1)
        parent = p.getparent()
        idx = parent.index(p)
        parent.remove(p)

        # entsprechenden Abschnitt einfuegen, falls vorhanden
        for b in sections.get(num, []):
            elem = getattr(b, '_element', b)
            parent.insert(idx, deepcopy(elem))
            idx += 1

        # ggf. Fallback-Icon einfuegen

            continue
        num = m.group(1)
        parent = p.getparent()
        idx = parent.index(p)
        parent.remove(p)

        # entsprechenden Abschnitt einfuegen, falls vorhanden
        for b in sections.get(num, []):
            elem = getattr(b, '_element', b)
            parent.insert(idx, deepcopy(elem))
            idx += 1

        # ggf. Fallback-Icon einfuegen

    # Regex zum Finden von Platzhaltern wie {SECTION_1} oder {{SECTION_1}}
    # (manche Templates nutzen doppelte geschweifte Klammern)
    pattern = re.compile(r"\{{1,2}SECTION_(\d+)\}{1,2}")
    # Regex zum Finden von Platzhaltern wie {SECTION_1}
    pattern = re.compile(r"{SECTION_(\d+)}")



    for block in list(iter_block_items(tpl)):
        if not isinstance(block, Paragraph):
            continue
        text = block.text
        m = pattern.search(text)
        if not m:
            continue
        num = m.group(1)
        idx = body.index(block._element)
        body.remove(block._element)

        # entsprechenden Abschnitt einfügen, falls vorhanden
        for b in sections.get(num, []):
            elem = getattr(b, '_element', b)
            body.insert(idx, deepcopy(elem))
            idx += 1

        # ggf. Fallback-Icon einfügen

        icon = os.path.join(ICONS_DIR, f'GHS{num}.png')
        if os.path.isfile(icon):
            w_in, _ = get_image_size_inches(icon)
            pic_p = tpl.add_paragraph()
            run = pic_p.add_run()
            run.add_picture(icon, width=Inches(w_in))

            if pic_p._p.getparent() is not None:
                pic_p._p.getparent().remove(pic_p._p)
            parent.insert(idx, pic_p._p)


            body.remove(pic_p._p)
            parent.insert(idx, pic_p._p)

            body.insert(idx, pic_p._p)

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
