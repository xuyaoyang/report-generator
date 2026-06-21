"""Render a DOCX to page PNGs using Microsoft Word and PyMuPDF.

This is a Windows-friendly fallback for layout QA when LibreOffice/Poppler is
not installed. It uses Word COM to export the DOCX to PDF, then rasterizes the
PDF with PyMuPDF.
"""
from __future__ import annotations

import argparse
from pathlib import Path
import tempfile

import fitz
import win32com.client


WD_EXPORT_FORMAT_PDF = 17
WD_ALERTS_NONE = 0


def export_pdf_with_word(docx_path: Path, pdf_path: Path) -> None:
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = WD_ALERTS_NONE
    doc = None
    try:
        doc = word.Documents.Open(str(docx_path.resolve()), ReadOnly=True)
        doc.ExportAsFixedFormat(str(pdf_path.resolve()), WD_EXPORT_FORMAT_PDF)
    finally:
        if doc is not None:
            doc.Close(False)
        word.Quit()


def rasterize_pdf(pdf_path: Path, output_dir: Path, dpi: int) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    scale = dpi / 72.0
    matrix = fitz.Matrix(scale, scale)
    result = []
    with fitz.open(pdf_path) as pdf:
        for index, page in enumerate(pdf, start=1):
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            out = output_dir / f"page-{index:03d}.png"
            pix.save(out)
            result.append(out)
    return result


def render_docx(docx_path: Path, output_dir: Path, dpi: int, emit_pdf: bool) -> list[Path]:
    docx_path = docx_path.resolve()
    output_dir = output_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    if emit_pdf:
        pdf_path = output_dir / f"{docx_path.stem}.pdf"
        export_pdf_with_word(docx_path, pdf_path)
        return rasterize_pdf(pdf_path, output_dir, dpi)

    with tempfile.TemporaryDirectory(prefix="word_render_") as tmp:
        pdf_path = Path(tmp) / f"{docx_path.stem}.pdf"
        export_pdf_with_word(docx_path, pdf_path)
        return rasterize_pdf(pdf_path, output_dir, dpi)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx", help="Path to the DOCX file to render.")
    parser.add_argument("--output-dir", required=True, help="Directory for page PNGs.")
    parser.add_argument("--dpi", type=int, default=150)
    parser.add_argument("--emit-pdf", action="store_true")
    args = parser.parse_args()

    pages = render_docx(Path(args.docx), Path(args.output_dir), args.dpi, args.emit_pdf)
    print(f"Rendered {len(pages)} page(s) to {Path(args.output_dir).resolve()}")
    for page in pages:
        print(page)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
