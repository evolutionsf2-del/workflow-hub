import os
import logging
import openpyxl
from openpyxl.utils import get_column_letter
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

logger = logging.getLogger(__name__)

def set_cell_background(cell, fill_hex="2563EB"):
    """Sets cell background color in docx table."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Sets cell padding (in twentieths of a point / dxa)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def convert_excel_to_docx(excel_path: str, docx_path: str, selected_sheets: list[str] = None) -> bool:
    """
    Converts an Excel (.xlsx / .xls) workbook to a Word (.docx) document
    with nicely styled tables and headers.
    """
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        doc = Document()

        # Adjust page margins (narrow margins for better table space)
        for section in doc.sections:
            section.top_margin = Inches(0.7)
            section.bottom_margin = Inches(0.7)
            section.left_margin = Inches(0.7)
            section.right_margin = Inches(0.7)

        # Main Document Title
        title_p = doc.add_paragraph()
        title_run = title_p.add_run(f"เอกสารรายงานจาก Excel: {os.path.basename(excel_path)}")
        title_run.font.size = Pt(16)
        title_run.font.bold = True
        title_run.font.color.rgb = RGBColor(30, 58, 138) # Dark Navy
        title_p.paragraph_format.space_after = Pt(14)

        sheets_to_process = selected_sheets if selected_sheets else wb.sheetnames

        for sheet_idx, sheet_name in enumerate(sheets_to_process):
            if sheet_name not in wb.sheetnames:
                continue

            ws = wb[sheet_name]

            # Sheet Title Heading
            if sheet_idx > 0:
                doc.add_page_break()

            h = doc.add_heading(level=2)
            h_run = h.add_run(f"แผ่นงาน: {sheet_name}")
            h_run.font.size = Pt(13)
            h_run.font.bold = True
            h_run.font.color.rgb = RGBColor(37, 99, 235) # Blue
            h.paragraph_format.space_before = Pt(10)
            h.paragraph_format.space_after = Pt(8)

            # Extract non-empty rows
            all_rows = list(ws.iter_rows(values_only=True))

            # Filter out trailing completely empty rows
            while all_rows and all(v is None or str(v).strip() == "" for v in all_rows[-1]):
                all_rows.pop()

            if not all_rows:
                p = doc.add_paragraph("(ไม่มีข้อมูลในแผ่นงานนี้)")
                p.paragraph_format.italic = True
                continue

            # Determine max cols with data
            max_cols = max(len([c for c in row if c is not None]) for row in all_rows)
            # Find the actual used width
            max_col_index = 0
            for row in all_rows:
                for idx, cell_val in enumerate(row):
                    if cell_val is not None and str(cell_val).strip() != "":
                        if idx + 1 > max_col_index:
                            max_col_index = idx + 1

            if max_col_index == 0:
                p = doc.add_paragraph("(ไม่มีข้อมูลในแผ่นงานนี้)")
                p.paragraph_format.italic = True
                continue

            num_cols = max_col_index
            num_rows = len(all_rows)

            # Create Word Table
            table = doc.add_table(rows=num_rows, cols=num_cols)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.autofit = True

            # Style Table
            try:
                table.style = 'Table Grid'
            except Exception:
                pass

            for row_idx, row_data in enumerate(all_rows):
                word_row = table.rows[row_idx]
                is_header = (row_idx == 0)

                # Set canSplit and repeatHeader for header
                trPr = word_row._tr.get_or_add_trPr()
                trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
                if is_header:
                    trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

                for col_idx in range(num_cols):
                    cell = word_row.cells[col_idx]
                    val = row_data[col_idx] if col_idx < len(row_data) else ""
                    val_str = "" if val is None else str(val)

                    cell.text = val_str
                    p = cell.paragraphs[0]
                    p.paragraph_format.space_before = Pt(2)
                    p.paragraph_format.space_after = Pt(2)

                    if is_header:
                        set_cell_background(cell, fill_hex="2563EB")
                        set_cell_margins(cell, top=140, bottom=140, left=150, right=150)
                        if p.runs:
                            p.runs[0].font.bold = True
                            p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
                            p.runs[0].font.size = Pt(10)
                    else:
                        if row_idx % 2 == 1:
                            set_cell_background(cell, fill_hex="F8FAFC") # light zebra striping
                        set_cell_margins(cell, top=80, bottom=80, left=120, right=120)
                        if p.runs:
                            p.runs[0].font.size = Pt(9.5)
                            p.runs[0].font.color.rgb = RGBColor(30, 41, 59)

            doc.add_paragraph().paragraph_format.space_after = Pt(12)

        doc.save(docx_path)
        wb.close()
        return True
    except Exception as e:
        logger.error(f"Error converting Excel to DOCX: {e}")
        raise e
