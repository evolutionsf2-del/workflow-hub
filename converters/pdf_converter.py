import os
import gc
import logging
from typing import Optional, Tuple
import fitz
from pdf2docx import Converter

logger = logging.getLogger(__name__)

# Maximum pages allowed per conversion on Render Free Tier (512MB RAM)
MAX_PAGES_LIMIT = 15

def get_pdf_page_count(pdf_path: str) -> int:
    """
    Returns the total number of pages in the given PDF file.
    Uses PyMuPDF (fitz) for fast, low-memory reading.
    """
    doc = None
    try:
        doc = fitz.open(pdf_path)
        return len(doc)
    finally:
        if doc is not None:
            doc.close()
            del doc
        gc.collect()

def convert_pdf_to_docx(
    pdf_path: str,
    docx_path: str,
    start_page: Optional[int] = None,
    end_page: Optional[int] = None,
    max_pages: int = MAX_PAGES_LIMIT
) -> Tuple[bool, int, int, int]:
    """
    Converts a PDF file or a specific page range to Word (.docx) format while preserving layout.
    Enforces a strict page limit per conversion to prevent Out-Of-Memory (OOM) crashes on Render Free tier.

    Args:
        pdf_path: Path to the input PDF file.
        docx_path: Path where the converted DOCX file will be saved.
        start_page: 1-based start page index (optional).
        end_page: 1-based end page index (optional).
        max_pages: Maximum pages allowed per conversion (default: 15).

    Returns:
        Tuple of (success: bool, actual_start: int, actual_end: int, total_pages: int)
    """
    total_pages = get_pdf_page_count(pdf_path)
    if total_pages == 0:
        raise ValueError("ไฟล์ PDF นี้ไม่มีหน้าข้อมูลที่สามารถแปลงได้")

    # Determine 1-based page range
    is_custom_range = (start_page is not None or end_page is not None)
    p_start = int(start_page) if start_page is not None else 1
    p_end = int(end_page) if end_page is not None else total_pages

    # Validation
    if p_start < 1:
        raise ValueError(f"หน้าเริ่มต้นต้องไม่น้อยกว่า 1 (ระบุมา: {p_start})")
    if p_end > total_pages:
        raise ValueError(f"หน้าสิ้นสุด ({p_end}) ต้องไม่เกินจำนวนหน้าทั้งหมดของเอกสาร ({total_pages} หน้า)")
    if p_start > p_end:
        raise ValueError(f"หน้าเริ่มต้น ({p_start}) ต้องไม่มากกว่าหน้าสิ้นสุด ({p_end})")

    pages_to_convert = p_end - p_start + 1

    # Check page limit
    if pages_to_convert > max_pages:
        if not is_custom_range:
            raise ValueError(
                f"ไฟล์มีจำนวนหน้าทั้งหมด {total_pages} หน้า ซึ่งเกินขีดจำกัดความปลอดภัยของเซิร์ฟเวอร์ฟรี "
                f"(จำกัดไม่เกิน {max_pages} หน้าต่อการแปลง) กรุณาใช้ตัวเลือก 'กำหนดช่วงหน้า' เพื่อแบ่งแปลงทีละส่วน "
                f"(เช่น หน้า 1-{max_pages} และ {max_pages + 1}-{min(total_pages, 2 * max_pages)}) "
                f"หรือเปิดรันโปรแกรมบนคอมพิวเตอร์ของคุณ (Localhost)"
            )
        else:
            raise ValueError(
                f"ช่วงหน้าที่เลือกมีจำนวน {pages_to_convert} หน้า (หน้า {p_start} ถึง {p_end}) "
                f"ซึ่งเกินขีดจำกัดความปลอดภัยของเซิร์ฟเวอร์ฟรี (ไม่เกิน {max_pages} หน้าต่อการแปลง) "
                f"กรุณาระบุช่วงหน้าไม่เกิน {max_pages} หน้าต่อครั้ง"
            )

    cv = None
    src_doc = None
    sub_doc = None
    sub_pdf_path = None

    try:
        # If converting a subset of pages, extract them into a temporary mini-PDF first.
        # This dramatically cuts down memory usage because pdf2docx will never load unrelated pages into RAM.
        if pages_to_convert < total_pages:
            sub_pdf_path = f"{pdf_path}_sub_{p_start}_{p_end}.pdf"
            src_doc = fitz.open(pdf_path)
            sub_doc = fitz.open()
            # PyMuPDF uses 0-based indices: from_page to to_page (inclusive)
            sub_doc.insert_pdf(src_doc, from_page=p_start - 1, to_page=p_end - 1)
            sub_doc.save(sub_pdf_path)
            src_doc.close()
            sub_doc.close()
            src_doc = None
            sub_doc = None

            conversion_source = sub_pdf_path
        else:
            conversion_source = pdf_path

        # Convert using pdf2docx
        cv = Converter(conversion_source)
        cv.convert(docx_path)
        return True, p_start, p_end, total_pages

    except Exception as e:
        logger.error(f"Error converting PDF to DOCX: {e}")
        raise e
    finally:
        # Step 3: Cleanup and immediate memory reclamation
        if cv is not None:
            try:
                cv.close()
            except Exception:
                pass
            del cv

        if src_doc is not None:
            try:
                src_doc.close()
            except Exception:
                pass
            del src_doc

        if sub_doc is not None:
            try:
                sub_doc.close()
            except Exception:
                pass
            del sub_doc

        if sub_pdf_path and os.path.exists(sub_pdf_path):
            try:
                os.remove(sub_pdf_path)
            except Exception:
                pass

        # Force garbage collection immediately to return RAM to the OS
        gc.collect()
