import os
import sys
from pdf2docx import Converter
import re

# Standardpfade (werden im Container verwendet)
DEFAULT_INPUT_DIR = "/data/sample_pdfs"
DEFAULT_OUTPUT_DIR = "/data/output"

# Pfade ggf. überschreiben, damit das Skript auch lokal lauffähig ist
INPUT_DIR = os.environ.get("INPUT_DIR", DEFAULT_INPUT_DIR)
OUTPUT_DIR = os.environ.get("OUTPUT_DIR", DEFAULT_OUTPUT_DIR)

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
    print(f"Converted PDF: {pdf_path} -> {raw_docx_path}")


if __name__ == '__main__':
    os.makedirs(INPUT_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    for f in os.listdir(INPUT_DIR):
        if not f.lower().endswith('.pdf'):
            continue
        pdf_file = os.path.join(INPUT_DIR, f)
        base, _ = os.path.splitext(f)
        safe_base = safe_re.sub('_', base)
        output_docx = os.path.join(OUTPUT_DIR, f"{safe_base}.docx")

        try:
            print(f"Processing PDF: {f}...")
            pdf_to_raw_docx(pdf_file, output_docx)
        except Exception as exc:
            print(f"Error converting {f}: {exc}", file=sys.stderr)
            continue
