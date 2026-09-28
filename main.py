import os
import shutil
import uuid
from pathlib import Path
from typing import Optional, List

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware

from converters.pdf_converter import convert_pdf_to_docx
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
    """Safely removes a temporary file after response has been sent."""
    try:
        if path.exists():
            path.unlink()
    except Exception as e:
        print(f"Failed to remove file {path}: {e}")

@app.get("/", response_class=HTMLResponse)
async def serve_index():
    index_file = Path("static/index.html")
    if index_file.exists():
        return index_file.read_text(encoding="utf-8")
    return "<h1>Office Utility Hub is running!</h1>"

@app.get("/api/health")
async def health_check():
    return {"status": "ok", "message": "Office Utility Hub is ready"}

@app.post("/api/convert/pdf-to-word")
async def api_pdf_to_word(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...)
):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="กรุณาอัปโหลดไฟล์ที่มีนามสกุล .pdf เท่านั้น")

    file_id = str(uuid.uuid4())
    input_filename = f"{file_id}_{file.filename}"
    input_path = TEMP_DIR / input_filename
    
    base_name = os.path.splitext(file.filename)[0]
    output_filename = f"{base_name}.docx"
    output_path = TEMP_DIR / f"{file_id}_{output_filename}"

    try:
        # Save uploaded file
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Convert
        convert_pdf_to_docx(str(input_path), str(output_path))

        # Schedule cleanup
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
        raise HTTPException(status_code=500, detail=f"เกิดข้อผิดพลาดในการแปลงไฟล์: {str(e)}")

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

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
