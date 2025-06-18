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
    # yield paragraphs and tables
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
        if text.upper().startswith('ABSCHNITT'):
            parts = text.split()
            if len(parts) >= 2:
                num = ''.join(ch for ch in parts[1] if ch.isdigit())
                current = num
                sections[num] = []
        if current:
            sections[current].append(block)
    return sections


def merge_into_template(sections, template_path, out_path):
    if not os.path.isfile(template_path):
        print(f"Template not found: {template_path}")
        return
    tpl = Document(template_path)
    body = tpl.element.body
    # for each block (p or table) in template
    for block in list(iter_block_items(tpl)):
        if not isinstance(block, Paragraph):
            continue
        text = block.text
        for num, blocks in sections.items():
            placeholder = f'{{{{SECTION_{num}}}}}'
            if placeholder in text:
                idx = body.index(block._element)
                body.remove(block._element)
                # insert raw content
                for b in blocks:
                    elem = getattr(b, '_element', b)
                    body.insert(idx, elem)
                    idx += 1
                # fallback icon
                icon = os.path.join(ICONS_DIR, f'GHS{num}.png')
                if os.path.isfile(icon):
                    w_in, _ = get_image_size_inches(icon)
                    pic_p = tpl.add_paragraph()
                    run = pic_p.add_run()
                    run.add_picture(icon, width=Inches(w_in))
                    body.insert(idx, pic_p._p)
                break
    tpl.save(out_path)


if __name__ == '__main__':
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(ICONS_DIR, exist_ok=True)
    for f in os.listdir(INPUT_DIR):
        if not f.lower().endswith('.pdf'):
            continue
        pdf = os.path.join(INPUT_DIR, f)
        raw = os.path.join(OUTPUT_DIR, f.replace('.pdf', '_raw.docx'))
        final = os.path.join(OUTPUT_DIR, f.replace('.pdf', '.docx'))
        print('Processing', f)
        pdf_to_raw_docx(pdf, raw)
        secs = extract_sections(raw)
        merge_into_template(secs, TEMPLATE_PATH, final)
        print('Saved', final)
