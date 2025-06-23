import os
import sys
import re
from typing import Dict, List
from pdf2docx import Converter
from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph
from copy import deepcopy

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

# Regex für Abschnitts-Header (flexibel für Deutsch/Englisch)
header_re = re.compile(r"^\s*(?:Abschnitt|Section)\s*\.?\s*(\d+)\s*[:\.-]?", re.I)

def pdf_to_raw_docx(pdf_path: str, raw_docx_path: str) -> None:
    """
    Konvertiert PDF direkt in DOCX mit pdf2docx.
    """
    os.makedirs(os.path.dirname(raw_docx_path), exist_ok=True)
    cv = Converter(pdf_path)
    cv.convert(raw_docx_path, start=0, end=None)
    cv.close()


def iter_block_items(parent):
    # yield paragraphs and tables only
    for child in parent.element.body:
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def extract_sections(raw_docx_path: str) -> Dict[str, List]:
    """
    Teilt ein rohes DOCX in Abschnitte auf, die durch 'Abschnitt X' oder 'Section X' markiert werden.
    """
    doc = Document(raw_docx_path)
    sections: Dict[str, List] = {}
    current: str = None

    for block in iter_block_items(doc):
        # Paragraph-Überschrift
        if isinstance(block, Paragraph):
            text = block.text.strip()
            m = header_re.match(text)
            if m:
                current = m.group(1)
                sections[current] = []
                continue
            if current:
                sections[current].append(block)

        # Tabellenblock
        elif isinstance(block, Table):
            # Suche nach einer Abschnitts-Überschrift innerhalb der Tabelle
            found_header = False
            header_row_idx = None
            header_num = None
            for ri, row in enumerate(block.rows):
                for cell in row.cells:
                    for para in cell.paragraphs:
                        m = header_re.match(para.text.strip())
                        if m:
                            header_num = m.group(1)
                            found_header = True
                            header_row_idx = ri
                            break
                    if found_header:
                        break
                if found_header:
                    break

            if found_header:
                current = header_num
                sections[current] = []
                # Kopiere Tabelle ohne Überschriftszeilen
                tbl_elem = deepcopy(block._element)
                # Entferne alle Zeilen bis header_row_idx
                for _ in range(header_row_idx + 1):
                    tbl_elem.tr_lst.pop(0)
                if tbl_elem.tr_lst:
                    sections[current].append(Table(tbl_elem, doc))
            else:
                # Ansonsten gesamte Tabelle zum aktuellen Abschnitt hinzufügen
                if current:
                    sections[current].append(block)

    return sections


def merge_into_template(sections: Dict[str, List], template_path: str, out_path: str) -> None:
    print(f">>> merge_into_template lädt Template von: {template_path}")
    if not os.path.isfile(template_path):
        print(f"Template not found: {template_path}", file=sys.stderr)
        return

    tpl = Document(template_path)
    body = tpl.element.body
    # Platzhalter-Muster: {SECTION_1} oder {SECTION 1}
    pattern = re.compile(r"\{\s*SECTION[_ ]?(\d+)\s*\}", re.I)

    # Alle Blöcke im Template durchgehen
    for block in list(iter_block_items(tpl)):
        if isinstance(block, Paragraph):
            text = block.text
            m = pattern.search(text)
            if not m:
                continue
            num = m.group(1)
            idx = body.index(block._element)
            # Platzhalter-Paragraph entfernen
            body.remove(block._element)
            # Sektionselemente einfügen
            for elem in sections.get(num, []):
                e = getattr(elem, '_element', elem)
                body.insert(idx, deepcopy(e))
                idx += 1

    tpl.save(out_path)


if __name__ == '__main__':
    os.makedirs(INPUT_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(ICONS_DIR, exist_ok=True)

    for f in os.listdir(INPUT_DIR):
        if not f.lower().endswith('.pdf'):
            continue
        try:
            pdf_path = os.path.join(INPUT_DIR, f)
            base, _ = os.path.splitext(f)
            raw_docx = os.path.join(OUTPUT_DIR, f"{base}_raw.docx")
            final_docx = os.path.join(OUTPUT_DIR, f"{base}.docx")

            print(f"Processing {f}...")
            pdf_to_raw_docx(pdf_path, raw_docx)
            sections = extract_sections(raw_docx)
            merge_into_template(sections, TEMPLATE_PATH, final_docx)
            print(f"Saved {final_docx}")

            # Eingangs-PDF löschen, falls gewünscht
            if os.path.exists(pdf_path):
                os.remove(pdf_path)
                print(f"Removed input PDF: {pdf_path}")

        except Exception as e:
            print(f"Error processing file {f}: {e}", file=sys.stderr)
            continue
