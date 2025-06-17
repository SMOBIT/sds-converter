import sys
from pdf2docx import Converter
from docx import Document


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
    pdf, template_dir, out_dir = sys.argv[1], sys.argv[2], sys.argv[3]
    raw = f"{out_dir}/{pdf.split('/')[-1].replace('.pdf', '_raw.docx')}"
    final = f"{out_dir}/{pdf.split('/')[-1].replace('.pdf', '.docx')}"
    template = f"{template_dir}/master_template.docx"

    pdf_to_raw_docx(pdf, raw)
    secs = extract_sections(raw)
    merge_into_template(secs, template, final)
    print(f"Converted {pdf} → {final}")
