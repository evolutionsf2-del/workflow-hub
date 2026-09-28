import os
import gc
import shutil
import uuid
from pathlib import Path
from typing import Optional, List

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware

from converters.pdf_converter import convert_pdf_to_docx, get_pdf_page_count, MAX_PAGES_LIMIT
from converters.excel_converter import convert_excel_to_docx
import openpyxl

app = FastAPI(title="Office Utility Hub API", version="1.0.0")

# Enable CORS for flexibility
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

TEMP_DIR = Path("./temp_storage")
TEMP_DIR.mkdir(exist_ok=True)

# Mount static files
app.mount("/static", StaticFiles(directory="static"), name="static")

def cleanup_file(path: Path):
    """Safely removes a temporary file after response has been sent and frees memory."""
    try:
        if path.exists():
            path.unlink()
    except Exception as e:
        print(f"Failed to remove file {path}: {e}")
    finally:
        gc.collect()

@app.get("/", response_class=HTMLResponse)
async def serve_index():
    index_file = Path("static/index.html")
    if index_file.exists():
        return index_file.read_text(encoding="utf-8")
    return "<h1>Office Utility Hub is running!</h1>"

@app.get("/api/health")
async def health_check():
    return {"status": "ok", "message": "Office Utility Hub is ready"}

@app.post("/api/inspect/pdf-info")
async def inspect_pdf_info(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...)
):
    """Inspects PDF to return total page count and limits before conversion."""
    filename_lower = file.filename.lower()
    if not filename_lower.endswith(".pdf"):
        raise HTTPException(status_code=400, detail="กรุณาอัปโหลดไฟล์ PDF (.pdf)")

    file_id = str(uuid.uuid4())
    input_path = TEMP_DIR / f"{file_id}_{file.filename}"

    try:
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        total_pages = get_pdf_page_count(str(input_path))
        background_tasks.add_task(cleanup_file, input_path)
        return {
            "status": "success",
            "filename": file.filename,
            "total_pages": total_pages,
            "max_allowed_pages": MAX_PAGES_LIMIT
        }
    except Exception as e:
        cleanup_file(input_path)
        raise HTTPException(status_code=400, detail=f"ไม่สามารถตรวจสอบข้อมูลไฟล์ PDF ได้: {str(e)}")

@app.post("/api/convert/pdf-to-word")
async def api_pdf_to_word(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    start_page: Optional[int] = Form(None),
    end_page: Optional[int] = Form(None)
):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="กรุณาอัปโหลดไฟล์ที่มีนามสกุล .pdf เท่านั้น")

    file_id = str(uuid.uuid4())
    input_filename = f"{file_id}_{file.filename}"
    input_path = TEMP_DIR / input_filename
    
    base_name = os.path.splitext(file.filename)[0]
    temp_output_filename = f"{file_id}_output.docx"
    output_path = TEMP_DIR / temp_output_filename

    try:
        # Save uploaded file
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Convert with page limit & range check
        _, actual_start, actual_end, total_pages = convert_pdf_to_docx(
            pdf_path=str(input_path),
            docx_path=str(output_path),
            start_page=start_page,
            end_page=end_page,
            max_pages=MAX_PAGES_LIMIT
        )

        # Determine download filename
        if actual_start == 1 and actual_end == total_pages:
            download_filename = f"{base_name}.docx"
        else:
            download_filename = f"{base_name}_หน้า{actual_start}-{actual_end}.docx"

        # Schedule cleanup
        background_tasks.add_task(cleanup_file, input_path)
        background_tasks.add_task(cleanup_file, output_path)

        return FileResponse(
            path=output_path,
            filename=download_filename,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except ValueError as ve:
        cleanup_file(input_path)
        if output_path.exists():
            cleanup_file(output_path)
        raise HTTPException(status_code=400, detail=str(ve))
    except Exception as e:
        cleanup_file(input_path)
        if output_path.exists():
            cleanup_file(output_path)
        raise HTTPException(status_code=500, detail=f"เกิดข้อผิดพลาดในการแปลงไฟล์: {str(e)}")
    finally:
        gc.collect()

@app.post("/api/inspect/excel-sheets")
async def inspect_excel_sheets(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...)
):
    filename_lower = file.filename.lower()
    if not (filename_lower.endswith(".xlsx") or filename_lower.endswith(".xls")):
        raise HTTPException(status_code=400, detail="กรุณาอัปโหลดไฟล์ Excel (.xlsx หรือ .xls)")

    file_id = str(uuid.uuid4())
    input_path = TEMP_DIR / f"{file_id}_{file.filename}"

    try:
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        wb = openpyxl.load_workbook(input_path, read_only=True)
        sheets = wb.sheetnames
        wb.close()
        cleanup_file(input_path)
        return {"sheets": sheets}
    except Exception as e:
        cleanup_file(input_path)
        raise HTTPException(status_code=500, detail=f"ไม่สามารถอ่านชีตในไฟล์ได้: {str(e)}")

@app.post("/api/convert/excel-to-word")
async def api_excel_to_word(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    sheets: Optional[str] = Form(None)
):
    filename_lower = file.filename.lower()
    if not (filename_lower.endswith(".xlsx") or filename_lower.endswith(".xls")):
        raise HTTPException(status_code=400, detail="กรุณาอัปโหลดไฟล์ Excel (.xlsx หรือ .xls)")

    file_id = str(uuid.uuid4())
    input_filename = f"{file_id}_{file.filename}"
    input_path = TEMP_DIR / input_filename

    base_name = os.path.splitext(file.filename)[0]
    output_filename = f"{base_name}.docx"
    output_path = TEMP_DIR / f"{file_id}_{output_filename}"

    try:
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        selected_sheets = None
        if sheets and sheets.strip():
            selected_sheets = [s.strip() for s in sheets.split(",") if s.strip()]

        convert_excel_to_docx(str(input_path), str(output_path), selected_sheets=selected_sheets)

        background_tasks.add_task(cleanup_file, input_path)
        background_tasks.add_task(cleanup_file, output_path)

        return FileResponse(
            path=output_path,
            filename=output_filename,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        cleanup_file(input_path)
        cleanup_file(output_path)
        raise HTTPException(status_code=500, detail=f"เกิดข้อผิดพลาดในการแปลงไฟล์ Excel: {str(e)}")
    finally:
        gc.collect()

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
