"""
Convert PDF/DOC/DOCX material certificates to images.
Each source document → a subfolder of PNG images (one per page).
"""
import os
import sys
import re

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MATERIAL_DIR = os.path.join(BASE_DIR, '材质单')

# Category folder name → category key mapping
FOLDER_TO_CATEGORY = {
    '钢板材质单': '钢板',
    '螺纹钢筋材质单': '钢筋',
    '圆钢材质单': '圆钢',
    '铅芯': '铅',
    '油漆': '油漆',
    '商混材质单': '商混',
    '缸筒材质单': '缸筒',
    '螺栓材质单': '螺栓',
}

DPI = 150  # resolution for rendering


def sanitize(name):
    """Remove problematic characters from folder names."""
    name = os.path.splitext(name)[0]
    name = re.sub(r'[<>:"/\\|?*]', '_', name)
    return name.strip()


def convert_pdf(pdf_path, out_dir):
    """Convert a PDF to PNG images, one per page. Returns list of output paths."""
    import fitz
    doc = fitz.open(pdf_path)
    paths = []
    for page_num in range(len(doc)):
        page = doc[page_num]
        mat = fitz.Matrix(DPI / 72, DPI / 72)
        pix = page.get_pixmap(matrix=mat)
        out_path = os.path.join(out_dir, f'page_{page_num + 1:02d}.png')
        pix.save(out_path)
        paths.append(out_path)
    doc.close()
    return paths


def convert_docx(docx_path, out_dir):
    """Convert DOCX to images via Word COM → temp PDF → PNG. Returns list of output paths."""
    import fitz
    import tempfile
    import pythoncom
    from win32com import client as wc

    pythoncom.CoInitialize()
    try:
        word = wc.Dispatch('Word.Application')
        word.Visible = False
        doc = word.Documents.Open(docx_path)
        # Temporarily save as PDF
        fd, tmp_pdf = tempfile.mkstemp(suffix='.pdf')
        os.close(fd)
        doc.SaveAs2(tmp_pdf, FileFormat=17)  # 17 = PDF
        doc.Close()
        word.Quit()
    finally:
        pythoncom.CoUninitialize()

    paths = convert_pdf(tmp_pdf, out_dir)
    os.remove(tmp_pdf)
    return paths


def convert_doc(doc_path, out_dir):
    """Convert .doc (old format) same as .docx."""
    import fitz
    import tempfile
    import pythoncom
    from win32com import client as wc

    pythoncom.CoInitialize()
    try:
        word = wc.Dispatch('Word.Application')
        word.Visible = False
        doc = word.Documents.Open(doc_path)
        fd, tmp_pdf = tempfile.mkstemp(suffix='.pdf')
        os.close(fd)
        doc.SaveAs2(tmp_pdf, FileFormat=17)
        doc.Close()
        word.Quit()
    finally:
        pythoncom.CoUninitialize()

    paths = convert_pdf(tmp_pdf, out_dir)
    os.remove(tmp_pdf)
    return paths


EXT_HANDLERS = {
    '.pdf': convert_pdf,
    '.docx': convert_docx,
    '.doc': convert_doc,
}


def main():
    total = 0
    total_pages = 0
    skipped = 0
    errors = []

    for root, dirs, files in os.walk(MATERIAL_DIR):
        for fname in files:
            ext = os.path.splitext(fname)[1].lower()
            if ext not in EXT_HANDLERS:
                continue

            filepath = os.path.join(root, fname)
            safe_name = sanitize(fname)
            out_dir = os.path.join(root, safe_name)

            if os.path.exists(out_dir) and os.listdir(out_dir):
                skipped += 1
                continue

            os.makedirs(out_dir, exist_ok=True)

            try:
                handler = EXT_HANDLERS[ext]
                paths = handler(filepath, out_dir)
                total += 1
                total_pages += len(paths)
                print(f'  OK [{len(paths)}p] {fname}')
            except Exception as e:
                errors.append((fname, str(e)))
                print(f'  FAIL {fname}: {e}')

    print(f'\n=== Done ===')
    print(f'Converted: {total} files, {total_pages} pages total')
    print(f'Skipped (already converted): {skipped}')
    if errors:
        print(f'Errors ({len(errors)}):')
        for fname, err in errors:
            print(f'  - {fname}: {err}')


if __name__ == '__main__':
    main()
