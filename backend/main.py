from fastapi import FastAPI, UploadFile, File, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
import os
import shutil
import uuid
import subprocess
import mimetypes

from pdf2docx import Converter
from PIL import Image
from PyPDF2 import PdfMerger, PdfReader, PdfWriter

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

UPLOAD_DIR = "uploads"
OUTPUT_DIR = "outputs"

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


@app.get("/")
def home():
    return {"message": "Convertify Forge API is running"}


def save_upload(file: UploadFile) -> str:
    file_id = str(uuid.uuid4())
    safe_name = file.filename.replace(" ", "_")
    path = os.path.join(UPLOAD_DIR, f"{file_id}_{safe_name}")

    with open(path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    return path


def is_pdf(file: UploadFile) -> bool:
    return file.filename.lower().endswith(".pdf")


@app.post("/convert")
async def convert_file(
    conversion_type: str = Form(...),
    file: UploadFile = File(None),
    files: list[UploadFile] = File(None),
    start_page: int = Form(1),
    end_page: int = Form(1),
):
    try:
        if conversion_type == "docx_to_pdf":
            if not file:
                return JSONResponse({"error": "Please upload a DOCX file."}, status_code=400)

            if not file.filename.lower().endswith(".docx"):
                return JSONResponse({"error": "DOCX to PDF requires a .docx file."}, status_code=400)

            input_path = save_upload(file)

            subprocess.run(
                [
                    "soffice",
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    OUTPUT_DIR,
                    input_path,
                ],
                check=True,
            )

            base_name = os.path.splitext(os.path.basename(input_path))[0]
            output_path = os.path.join(OUTPUT_DIR, f"{base_name}.pdf")

            if not os.path.exists(output_path):
                return JSONResponse(
                    {"error": "DOCX to PDF conversion failed. Make sure LibreOffice is installed."},
                    status_code=500,
                )

            return FileResponse(output_path, filename="converted.pdf", media_type="application/pdf")

        elif conversion_type == "pdf_to_docx":
            if not file:
                return JSONResponse({"error": "Please upload a PDF file."}, status_code=400)

            if not is_pdf(file):
                return JSONResponse({"error": "PDF to DOCX requires a PDF file."}, status_code=400)

            input_path = save_upload(file)
            output_path = os.path.join(OUTPUT_DIR, f"{uuid.uuid4()}.docx")

            cv = Converter(input_path)
            cv.convert(output_path)
            cv.close()

            return FileResponse(
                output_path,
                filename="converted.docx",
                media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

        elif conversion_type == "image_to_pdf":
            if not file:
                return JSONResponse({"error": "Please upload an image file."}, status_code=400)

            input_path = save_upload(file)
            output_path = os.path.join(OUTPUT_DIR, f"{uuid.uuid4()}.pdf")

            image = Image.open(input_path)
            if image.mode in ("RGBA", "P"):
                image = image.convert("RGB")

            image.save(output_path, "PDF")

            return FileResponse(output_path, filename="image-converted.pdf", media_type="application/pdf")

        elif conversion_type == "merge_pdf":
            if not files or len(files) < 2:
                return JSONResponse(
                    {"error": "Merge PDF requires at least 2 PDF files."},
                    status_code=400,
                )

            merger = PdfMerger()

            for uploaded_file in files:
                if not is_pdf(uploaded_file):
                    return JSONResponse(
                        {"error": "All merged files must be PDF files."},
                        status_code=400,
                    )

                input_path = save_upload(uploaded_file)
                merger.append(input_path)

            output_path = os.path.join(OUTPUT_DIR, f"{uuid.uuid4()}_merged.pdf")
            merger.write(output_path)
            merger.close()

            return FileResponse(output_path, filename="merged.pdf", media_type="application/pdf")

        elif conversion_type == "split_pdf":
            if not file:
                return JSONResponse({"error": "Please upload a PDF file."}, status_code=400)

            if not is_pdf(file):
                return JSONResponse({"error": "Split PDF requires a PDF file."}, status_code=400)

            input_path = save_upload(file)
            reader = PdfReader(input_path)
            total_pages = len(reader.pages)

            if start_page < 1 or end_page < 1:
                return JSONResponse({"error": "Page numbers must be 1 or higher."}, status_code=400)

            if start_page > end_page:
                return JSONResponse({"error": "Start page cannot be greater than end page."}, status_code=400)

            if end_page > total_pages:
                return JSONResponse(
                    {"error": f"This PDF only has {total_pages} pages."},
                    status_code=400,
                )

            writer = PdfWriter()

            for page_number in range(start_page - 1, end_page):
                writer.add_page(reader.pages[page_number])

            output_path = os.path.join(OUTPUT_DIR, f"{uuid.uuid4()}_split.pdf")

            with open(output_path, "wb") as output_file:
                writer.write(output_file)

            return FileResponse(
                output_path,
                filename=f"pages-{start_page}-to-{end_page}.pdf",
                media_type="application/pdf",
            )

        elif conversion_type == "compress_pdf":
            if not file:
                return JSONResponse({"error": "Please upload a PDF file."}, status_code=400)

            if not is_pdf(file):
                return JSONResponse({"error": "Compress PDF requires a PDF file."}, status_code=400)

            input_path = save_upload(file)
            output_path = os.path.join(OUTPUT_DIR, f"{uuid.uuid4()}_compressed.pdf")

            reader = PdfReader(input_path)
            writer = PdfWriter()

            for page in reader.pages:
                page.compress_content_streams()
                writer.add_page(page)

            with open(output_path, "wb") as output_file:
                writer.write(output_file)

            return FileResponse(output_path, filename="compressed.pdf", media_type="application/pdf")

        elif conversion_type == "detect_format":
            if not file:
                return JSONResponse({"error": "Please upload a file."}, status_code=400)

            input_path = save_upload(file)
            file_size = os.path.getsize(input_path)
            extension = os.path.splitext(file.filename)[1]
            mime_type = mimetypes.guess_type(file.filename)[0]

            return {
                "filename": file.filename,
                "extension": extension,
                "mime_type": mime_type or "unknown",
                "size_bytes": file_size,
                "size_kb": round(file_size / 1024, 2),
            }

        return JSONResponse({"error": f"Unsupported conversion type: {conversion_type}"}, status_code=400)

    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)


@app.post("/merge")
async def merge_pdfs(files: list[UploadFile] = File(...)):
    if len(files) < 2:
        return JSONResponse({"error": "Please upload at least 2 PDF files."}, status_code=400)

    merger = PdfMerger()

    for uploaded_file in files:
        if not is_pdf(uploaded_file):
            return JSONResponse({"error": "All merged files must be PDF files."}, status_code=400)

        input_path = save_upload(uploaded_file)
        merger.append(input_path)

    output_path = os.path.join(OUTPUT_DIR, f"{uuid.uuid4()}_merged.pdf")
    merger.write(output_path)
    merger.close()

    return FileResponse(output_path, filename="merged.pdf", media_type="application/pdf")