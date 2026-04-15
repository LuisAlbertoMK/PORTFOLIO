#!/usr/bin/env python3
import sys
from pathlib import Path

try:
    from docx import Document
except ImportError:
    print("ERROR: python-docx no instalado")
    print("Ejecutar: pip install python-docx")
    sys.exit(1)


def extract_text_from_docx(docx_path):
    doc = Document(docx_path)
    paragraphs = [para.text.strip() for para in doc.paragraphs if para.text.strip()]
    return "\n\n".join(paragraphs)


def main():
    cv_files = [
        ("CV_EMPRESARIAL.docx", "CV_EMPRESARIAL.md"),
        ("CV_TECNICO.docx", "CV_TECNICO.md"),
    ]

    base_path = Path("/media/mk/MK/Desarrollo/PORTFOLIO")

    for docx_name, md_name in cv_files:
        docx_path = base_path / docx_name
        md_path = base_path / md_name

        if not docx_path.exists():
            print(f"❌ No encontrado: {docx_path}")
            continue

        try:
            text = extract_text_from_docx(docx_path)
            md_path.write_text(text, encoding="utf-8")
            print(f"✓ Extraído: {docx_name} → {md_name}")
        except Exception as e:
            print(f"❌ Error con {docx_name}: {e}")


if __name__ == "__main__":
    main()
