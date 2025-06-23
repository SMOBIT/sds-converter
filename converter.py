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
from lxml import etree

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

print(f">>> Verwende TEMPLATE_PATH: {TEMPLATE_PATH}")

# Regex für Abschnitts-Header (flexibel für Deutsch/Englisch)
# Beispiele: "Abschnitt 1:", "• Abschnitt 1:", "Section 2 –", "1. Abschnitt 3- Inhalt"
header_re = re.compile(
    r"^\s*(?:[\d•*\-–.]+\s*)?(?:Abschnitt|Section)\s*\.?\s*(\d+)\s*[:.\-–]?",
    re.I,
)
# Regex für sichere Dateinamen
safe_re = re.compile(r"[^0-9A-Za-z]+")


def pdf_to_raw_docx(pdf_path: str, raw_docx_path: str) -> None:
    """
    Konvertiert PDF direkt in DOCX mit pdf2docx.
    """
    os.makedirs(os.path.dirname(raw_docx_path), exist_ok=True)
    cv = Converter(pdf_path)
    cv.convert(raw_docx_path, start=0, end=None)
    cv.close()


def extract_sections(raw_docx_path: str) -> Dict[str, List]:
    """
    Teilt ein rohes DOCX in Abschnitte auf, markiert durch 'Abschnitt X' oder 'Section X'.
    Sucht in allen Paragraph- und Tabellen-Knoten via XPath, um auch Shapes/Textfelder abzudecken.
    """
    doc = Document(raw_docx_path)
    sections: Dict[str, List] = {}
    current: str = None

    # Namespace für XPath
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

    # Paragraphs via XPath
    for p_elem in doc.element.body.xpath('.//w:p', namespaces=ns):
        texts = [t.text for t in p_elem.xpath('.//w:t', namespaces=ns) if t.text]
        full = ''.join(texts).strip()
        m = header_re.match(full)
        if m:
            current = m.group(1)
            sections.setdefault(current, [])
            continue
        if current and full:
            para = Paragraph(p_elem, doc)
            sections[current].append(para)

    # Tabellen via XPath
    for tbl_elem in doc.element.body.xpath('.//w:tbl', namespaces=ns):
        rows = tbl_elem.xpath('.//w:tr', namespaces=ns)
        header_sec = None
        header_row = None
        for ri, tr in enumerate(rows):
            texts = [t.text for t in tr.xpath('.//w:t', namespaces=ns) if t.text]
            full = ''.join(texts).strip()
            m = header_re.match(full)
            if m:
                header_sec = m.group(1)
                header_row = ri
                break
        if header_sec:
            current = header_sec
            sections.setdefault(current, [])
            clone = deepcopy(tbl_elem)
            clone_rows = clone.xpath('.//w:tr', namespaces=ns)
            # Entferne Header-Zeile
            clone.remove(clone_rows[header_row])
            # Verbleibende Zeilen prüfen
            if len(clone_rows) > header_row + 1:
                table = Table(clone, doc)
                sections[current].append(table)
            else:
                # Nur Kopf vorhanden
                sections[current].append(Table(tbl_elem, doc))
        else:
            if current:
                sections[current].append(Table(tbl_elem, doc))

    return sections


def merge_into_template(sections: Dict[str, List], template_path: str, out_path: str) -> None:
    """
    Fügt die extrahierten Abschnitte in das Template an den Platzhaltern {SECTION X} ein,
    sucht dabei in allen w:p-Elementen via XPath.
    """
    print(f">>> merge_into_template lädt Template von: {template_path}")
    if not os.path.isfile(template_path):
        print(f"Template nicht gefunden: {template_path}", file=sys.stderr)
        return

    tpl = Document(template_path)
    pattern = re.compile(r"\{\s*SECTION[_ ]?(\d+)\s*\}", re.I)
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

    for p_elem in tpl.element.body.xpath('.//w:p', namespaces=ns):
        texts = [t.text for t in p_elem.xpath('.//w:t', namespaces=ns) if t.text]
        full_text = ''.join(texts)
        m = pattern.search(full_text)
        if not m:
            continue
        sec = m.group(1)
        parent = p_elem.getparent()
        idx = parent.index(p_elem)
        parent.remove(p_elem)
        for elem in sections.get(sec, []):
            new_elm = deepcopy(elem._element)
            parent.insert(idx, new_elm)
            idx += 1

    tpl.save(out_path)
    print(f">>> Template gespeichert: {out_path}")


if __name__ == '__main__':
    os.makedirs(INPUT_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(ICONS_DIR, exist_ok=True)

    for f in os.listdir(INPUT_DIR):
        if not f.lower().endswith('.pdf'):
            continue
        try:
            pdf_file = os.path.join(INPUT_DIR, f)
            base, _ = os.path.splitext(f)
            safe_base = safe_re.sub('_', base)
            raw = os.path.join(OUTPUT_DIR, f"{safe_base}_raw.docx")
            final = os.path.join(OUTPUT_DIR, f"{safe_base}.docx")

            print(f"Processing {f}...")
            pdf_to_raw_docx(pdf_file, raw)
            sections = extract_sections(raw)
            print(f">>> Debug: Sections für {f}: {list(sections.keys())}")
            for s, els in sections.items(): print(f"  Sektion {s}: {len(els)} Elemente")

            merge_into_template(sections, TEMPLATE_PATH, final)
            print(f"Saved {final}")

            if os.path.exists(pdf_file):
                os.remove(pdf_file)
                print(f"Removed input PDF: {pdf_file}")

        except Exception as exc:
            print(f"Fehler bei {f}: {exc}", file=sys.stderr)
            continue
