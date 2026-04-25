import io
import os
from pathlib import Path
from typing import List, Optional

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from starlette.middleware.cors import CORSMiddleware

from pdf_utils import (
    add_page_numbers,
    add_watermark,
    censor_pdf,
    compare_pdfs,
    compress_pdf,
    crop_pdf,
    delete_pdf_pages,
    excel_to_pdf,
    extract_pdf_pages,
    extract_text_from_pdf,
    html_to_pdf,
    image_to_pdf,
    merge_pdfs,
    pdf_to_excel,
    pdf_to_images,
    pdf_to_powerpoint,
    pdf_to_word,
    powerpoint_to_pdf,
    protect_pdf,
    repair_pdf,
    reorder_pdf_pages,
    rotate_pdf,
    split_pdf,
    unlock_pdf,
    word_to_pdf,
)

app = FastAPI(
    title="NOVA PDF Tools",
    description="Prototype open-source d'un editeur PDF web.",
)

cors_origins = [
    origin.strip()
    for origin in os.getenv("NOVA_CORS_ORIGINS", "http://localhost:8000,http://127.0.0.1:8000").split(",")
    if origin.strip()
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=cors_origins,
    allow_methods=["*"],
    allow_headers=["*"],
)

static_dir = Path(__file__).resolve().parent / "app" / "static"
app.mount("/static", StaticFiles(directory=static_dir), name="static")


def file_download(content: bytes, media_type: str, filename: str) -> StreamingResponse:
    return StreamingResponse(
        io.BytesIO(content),
        media_type=media_type,
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.exception_handler(ValueError)
async def value_error_handler(_: Request, exc: ValueError):
    return JSONResponse(status_code=400, content={"detail": str(exc)})


@app.get("/", response_class=HTMLResponse)
async def index():
    content = (static_dir / "index.html").read_text(encoding="utf-8")
    return HTMLResponse(content=content)


@app.post("/api/merge")
async def api_merge(files: List[UploadFile] = File(...)):
    if len(files) < 2:
        raise HTTPException(status_code=400, detail="Au moins deux fichiers PDF sont necessaires pour fusionner.")
    return file_download(merge_pdfs(files), "application/pdf", "merged.pdf")


@app.post("/api/split")
async def api_split(file: UploadFile = File(...), pages: str = Form(...)):
    return file_download(split_pdf(file, pages), "application/pdf", "splitted.pdf")


@app.post("/api/reorder")
async def api_reorder(file: UploadFile = File(...), pages: str = Form(...)):
    return file_download(reorder_pdf_pages(file, pages), "application/pdf", "reordered.pdf")


@app.post("/api/rotate")
async def api_rotate(file: UploadFile = File(...), angle: int = Form(...), pages: Optional[str] = Form(None)):
    if angle % 90 != 0:
        raise HTTPException(status_code=400, detail="L'angle doit etre un multiple de 90.")
    return file_download(rotate_pdf(file, angle, pages), "application/pdf", "rotated.pdf")


@app.post("/api/crop")
async def api_crop(
    file: UploadFile = File(...),
    top: float = Form(0),
    right: float = Form(0),
    bottom: float = Form(0),
    left: float = Form(0),
):
    return file_download(crop_pdf(file, top, right, bottom, left), "application/pdf", "cropped.pdf")


@app.post("/api/compress")
async def api_compress(file: UploadFile = File(...)):
    return file_download(compress_pdf(file), "application/pdf", "compressed.pdf")


@app.post("/api/repair")
async def api_repair(file: UploadFile = File(...)):
    return file_download(repair_pdf(file), "application/pdf", "repaired.pdf")


@app.post("/api/convert/image-to-pdf")
async def api_image_to_pdf(file: UploadFile = File(...)):
    return file_download(image_to_pdf(file), "application/pdf", "converted.pdf")


@app.post("/api/convert/html-to-pdf")
async def api_html_to_pdf(file: UploadFile = File(...)):
    return file_download(html_to_pdf(file), "application/pdf", "html-converted.pdf")


@app.post("/api/convert/word-to-pdf")
async def api_word_to_pdf(file: UploadFile = File(...)):
    return file_download(word_to_pdf(file), "application/pdf", "word-converted.pdf")


@app.post("/api/convert/excel-to-pdf")
async def api_excel_to_pdf(file: UploadFile = File(...)):
    return file_download(excel_to_pdf(file), "application/pdf", "excel-converted.pdf")


@app.post("/api/convert/powerpoint-to-pdf")
async def api_powerpoint_to_pdf(file: UploadFile = File(...)):
    return file_download(powerpoint_to_pdf(file), "application/pdf", "powerpoint-converted.pdf")


@app.post("/api/delete")
async def api_delete(file: UploadFile = File(...), pages: str = Form(...)):
    return file_download(delete_pdf_pages(file, pages), "application/pdf", "deleted.pdf")


@app.post("/api/extract")
async def api_extract(file: UploadFile = File(...), pages: str = Form(...)):
    return file_download(extract_pdf_pages(file, pages), "application/pdf", "extracted.pdf")


@app.post("/api/ocr")
async def api_ocr(file: UploadFile = File(...)):
    return file_download(extract_text_from_pdf(file), "text/plain; charset=utf-8", "ocr.txt")


@app.post("/api/watermark")
async def api_watermark(file: UploadFile = File(...), text: str = Form(...), opacity: float = Form(0.3)):
    return file_download(add_watermark(file, text, opacity), "application/pdf", "watermarked.pdf")


@app.post("/api/pdf-to-jpg")
async def api_pdf_to_jpg(file: UploadFile = File(...)):
    return file_download(pdf_to_images(file), "application/zip", "images.zip")


@app.post("/api/pdf-to-word")
async def api_pdf_to_word(file: UploadFile = File(...)):
    return file_download(
        pdf_to_word(file),
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "converted.docx",
    )


@app.post("/api/pdf-to-excel")
async def api_pdf_to_excel(file: UploadFile = File(...)):
    return file_download(
        pdf_to_excel(file),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "converted.xlsx",
    )


@app.post("/api/pdf-to-powerpoint")
async def api_pdf_to_powerpoint(file: UploadFile = File(...)):
    return file_download(
        pdf_to_powerpoint(file),
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "converted.pptx",
    )


@app.post("/api/numbering")
async def api_numbering(
    file: UploadFile = File(...),
    format_str: str = Form("{page}"),
    position: str = Form("bottom-right"),
):
    return file_download(add_page_numbers(file, format_str, position), "application/pdf", "numbered.pdf")


@app.post("/api/protect")
async def api_protect(
    file: UploadFile = File(...),
    user_password: str = Form(...),
    owner_password: Optional[str] = Form(None),
):
    return file_download(protect_pdf(file, user_password, owner_password), "application/pdf", "protected.pdf")


@app.post("/api/unlock")
async def api_unlock(file: UploadFile = File(...), password: str = Form(...)):
    return file_download(unlock_pdf(file, password), "application/pdf", "unlocked.pdf")


@app.post("/api/compare")
async def api_compare(file_a: UploadFile = File(...), file_b: UploadFile = File(...)):
    return file_download(compare_pdfs(file_a, file_b), "application/json", "compare-report.json")


@app.post("/api/censor")
async def api_censor(
    file: UploadFile = File(...),
    terms: str = Form(...),
    case_sensitive: bool = Form(False),
):
    return file_download(censor_pdf(file, terms, case_sensitive), "application/pdf", "censored.pdf")
