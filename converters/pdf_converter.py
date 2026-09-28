import os
import logging
from pdf2docx import Converter

logger = logging.getLogger(__name__)

def convert_pdf_to_docx(pdf_path: str, docx_path: str, start_page: int = 0, end_page: int = None) -> bool:
    """
    Converts a PDF file to a Word (.docx) file preserving formatting and layout.
    """
    cv = None
    try:
        cv = Converter(pdf_path)
        if end_page is not None:
            cv.convert(docx_path, start=start_page, end=end_page)
        else:
            cv.convert(docx_path, start=start_page)
        return True
    except Exception as e:
        logger.error(f"Error converting PDF to DOCX: {e}")
        raise e
    finally:
        if cv:
            cv.close()
