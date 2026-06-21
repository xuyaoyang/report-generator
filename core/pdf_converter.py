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

    if pdf_path is None:
        pdf_path = docx_path.replace('.docx', '.pdf')

    os.makedirs(os.path.dirname(pdf_path), exist_ok=True)

    import win32com.client
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    try:
        doc = word.Documents.Open(docx_path)
        doc.SaveAs2(pdf_path, FileFormat=17)  # 17 = wdFormatPDF
        doc.Close()
        return pdf_path
    finally:
        word.Quit()
