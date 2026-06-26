"""Convert .docx to PDF using Word COM on Windows."""
import os


def docx_to_pdf(docx_path, pdf_path=None):
    """
    Convert a .docx file to PDF using Word COM.
    If pdf_path is not specified, saves next to .docx with .pdf extension.
    Returns the PDF path.
    """
    if not os.path.exists(docx_path):
        raise FileNotFoundError(f'Document not found: {docx_path}')

    docx_path = os.path.abspath(docx_path)
    if pdf_path is None:
        pdf_path = os.path.splitext(docx_path)[0] + '.pdf'
    pdf_path = os.path.abspath(pdf_path)

    os.makedirs(os.path.dirname(pdf_path), exist_ok=True)

    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        word = win32com.client.DispatchEx('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False
        doc = word.Documents.Open(docx_path, ReadOnly=True)
        doc.SaveAs2(pdf_path, FileFormat=17)  # 17 = wdFormatPDF
        return pdf_path
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
