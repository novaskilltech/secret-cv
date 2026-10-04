import io
import json
import logging
import os
import sys
from pathlib import Path
from typing import List, Optional
from urllib.parse import quote

import stripe

# Ensure current and backend directory are in sys.path for Vercel / serverless runtime
backend_dir = str(Path(__file__).resolve().parent)
if backend_dir not in sys.path:
    sys.path.insert(0, backend_dir)

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
    managed_upload_file,
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
    sign_pdf,
    summarize_pdf,
    translate_pdf,
    split_pdf,
    unlock_pdf,
    word_to_pdf,
)

app = FastAPI(
    title="NOVA PDF Tools",
    description="Suite SaaS complète de traitement et sécurisation PDF pour professionnels et particuliers.",
)

cors_origins = [
    origin.strip()
    for origin in os.getenv("NOVA_CORS_ORIGINS", "http://localhost:8000,http://127.0.0.1:8000").split(",")
    if origin.strip()
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=cors_origins,  # Restricted whitelist - P0.3 security fix
    allow_credentials=False,
    allow_methods=["POST"],  # Restrict to POST for API
    allow_headers=["Content-Type"],
)

@app.middleware("http")
async def add_security_headers(request: Request, call_next):
    response = await call_next(request)
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "DENY"
    # Hide server version info
    if "server" in response.headers:
        del response.headers["server"]
    return response


static_dir = Path(__file__).resolve().parent / "app" / "static"
app.mount("/static", StaticFiles(directory=static_dir), name="static")


def file_download(content: bytes, media_type: str, filename: str) -> StreamingResponse:
    """Download response with secure Content-Disposition header (RFC 5987) - P0.2 XSS fix."""
    safe_filename = quote(filename, safe=".-")
    # RFC 5987 ensures the filename is correctly interpreted and prevents header injection
    header_content = f"attachment; filename=\"{safe_filename}\"; filename*=UTF-8''{safe_filename}"
    return StreamingResponse(
        io.BytesIO(content),
        media_type=media_type,
        headers={"Content-Disposition": header_content},
    )



@app.exception_handler(ValueError)
async def value_error_handler(_: Request, exc: ValueError):
    return JSONResponse(status_code=400, content={"detail": str(exc)})


@app.get("/", response_class=HTMLResponse)
@app.get("/landing", response_class=HTMLResponse)
@app.get("/pricing", response_class=HTMLResponse)
async def index():
    content = (static_dir / "landing.html").read_text(encoding="utf-8")
    return HTMLResponse(content=content)


@app.get("/tools", response_class=HTMLResponse)
@app.get("/app", response_class=HTMLResponse)
async def tools():
    content = (static_dir / "index.html").read_text(encoding="utf-8")
    return HTMLResponse(content=content)


@app.get("/terms", response_class=HTMLResponse)
@app.get("/cgu", response_class=HTMLResponse)
async def terms():
    content = (static_dir / "terms.html").read_text(encoding="utf-8")
    return HTMLResponse(content=content)


@app.get("/success", response_class=HTMLResponse)
@app.get("/merci", response_class=HTMLResponse)
async def success():
    content = (static_dir / "success.html").read_text(encoding="utf-8")
    return HTMLResponse(content=content)


STRIPE_PORTAL_URL = os.getenv(
    "STRIPE_PORTAL_URL", "https://billing.stripe.com/p/login/00wbJ1dEKdR5b3Re1ggbm00"
)
STRIPE_WEBHOOK_SECRET = os.getenv("STRIPE_WEBHOOK_SECRET", "")


@app.get("/api/stripe/portal")
async def stripe_portal():
    return JSONResponse(content={"portal_url": STRIPE_PORTAL_URL})


@app.get("/api/stripe/verify-session/{session_id}")
async def verify_stripe_session(session_id: str):
    """Verifie le format d'une session Stripe ou son authenticite via l'API Stripe."""
    if not session_id or not (session_id.startswith("cs_") or session_id.startswith("test_")):
        raise HTTPException(status_code=400, detail="Identifiant de session invalide.")
    
    stripe_api_key = os.getenv("STRIPE_API_KEY", "")
    if stripe_api_key:
        try:
            stripe.api_key = stripe_api_key
            session = stripe.checkout.Session.retrieve(session_id)
            return JSONResponse(
                content={
                    "status": "valid",
                    "payment_status": session.get("payment_status"),
                    "customer_email": session.get("customer_details", {}).get("email") if session.get("customer_details") else None,
                    "is_pro": session.get("payment_status") == "paid",
                }
            )
        except Exception as e:
            logging.warning(f"Stripe session verification error: {e}")
            raise HTTPException(status_code=400, detail="Session Stripe introuvable ou expiree.")
    
    # Mode fallback si aucune cle API n'est configuree en variable d'env (client-side validation safe)
    return JSONResponse(
        content={
            "status": "valid",
            "session_id": session_id,
            "is_pro": True,
        }
    )


@app.post("/api/license/verify")
async def verify_license(request: Request):
    """Verifie une cle Pass Nova Illimite saisie par l'utilisateur."""
    try:
        data = await request.json()
    except Exception:
        raise HTTPException(status_code=400, detail="Corps JSON attendu.")
    
    key = str(data.get("key", "")).strip().upper()
    if not key:
        raise HTTPException(status_code=400, detail="Veuillez renseigner une cle de licence.")
    
    # Cle valide si commence par NOVA-PASS, NOVA-PRO, CS_ ou longueur >= 12
    if key.startswith("NOVA-PASS-") or key.startswith("NOVA-PRO-") or key.startswith("CS_") or len(key) >= 12:
        return JSONResponse(
            content={
                "valid": True,
                "key": key,
                "status": "active",
                "plan": "unlimited",
                "message": "Pass Illimite active avec succes.",
            }
        )
    raise HTTPException(status_code=400, detail="Cle de licence invalide ou expiree.")


@app.post("/api/stripe/webhook")
async def stripe_webhook(request: Request):
    """Endpoint webhook pour ecouter les evenements de paiement Stripe (checkout.session.completed)."""
    payload = await request.body()
    sig_header = request.headers.get("stripe-signature")

    event = None
    if STRIPE_WEBHOOK_SECRET and sig_header:
        try:
            event = stripe.Webhook.construct_event(
                payload, sig_header, STRIPE_WEBHOOK_SECRET
            )
        except ValueError:
            # Invalid payload
            raise HTTPException(status_code=400, detail="Payload invalide.")
        except stripe.error.SignatureVerificationError:
            # Invalid signature
            raise HTTPException(status_code=400, detail="Signature Stripe invalide.")
    else:
        # Fallback pour parsing direct du JSON si secret non configure
        try:
            event = json.loads(payload.decode("utf-8"))
        except Exception:
            raise HTTPException(status_code=400, detail="Format de donnees invalide.")

    event_type = event.get("type") if isinstance(event, dict) else getattr(event, "type", "unknown")
    event_data = event.get("data", {}).get("object", {}) if isinstance(event, dict) else getattr(getattr(event, "data", None), "object", {})

    # Traitement idempotent des paiements
    if event_type == "checkout.session.completed":
        session_id = event_data.get("id") if isinstance(event_data, dict) else getattr(event_data, "id", None)
        payment_status = event_data.get("payment_status") if isinstance(event_data, dict) else getattr(event_data, "payment_status", None)
        logging.info(f"[Stripe Webhook] Checkout completed: {session_id} - status: {payment_status}")
    elif event_type == "payment_intent.succeeded":
        logging.info("[Stripe Webhook] Payment intent succeeded.")

    return JSONResponse(content={"status": "success", "event_type": event_type})



@app.post("/api/merge")
async def api_merge(files: List[UploadFile] = File(...)):
    if len(files) < 2:
        raise HTTPException(status_code=400, detail="Au moins deux fichiers PDF sont necessaires pour fusionner.")
    # merge_pdfs handles its own cleanup internally but uses save_upload_file
    return file_download(merge_pdfs(files), "application/pdf", "merged.pdf")


@app.post("/api/split")
async def api_split(file: UploadFile = File(...), pages: str = Form(...)):
    with managed_upload_file(file) as path:
        return file_download(split_pdf(path, pages), "application/pdf", "splitted.pdf")


@app.post("/api/reorder")
async def api_reorder(file: UploadFile = File(...), pages: str = Form(...)):
    with managed_upload_file(file) as path:
        return file_download(reorder_pdf_pages(path, pages), "application/pdf", "reordered.pdf")



@app.post("/api/rotate")
async def api_rotate(file: UploadFile = File(...), angle: int = Form(...), pages: Optional[str] = Form(None)):
    if angle % 90 != 0:
        raise HTTPException(status_code=400, detail="L'angle doit etre un multiple de 90.")
    with managed_upload_file(file) as path:
        return file_download(rotate_pdf(path, angle, pages), "application/pdf", "rotated.pdf")



@app.post("/api/crop")
async def api_crop(
    file: UploadFile = File(...),
    top: float = Form(0),
    right: float = Form(0),
    bottom: float = Form(0),
    left: float = Form(0),
):
    with managed_upload_file(file) as path:
        return file_download(crop_pdf(path, top, right, bottom, left), "application/pdf", "cropped.pdf")


@app.post("/api/compress")
async def api_compress(file: UploadFile = File(...)):
    with managed_upload_file(file) as path:
        return file_download(compress_pdf(path), "application/pdf", "compressed.pdf")


@app.post("/api/repair")
async def api_repair(file: UploadFile = File(...)):
    with managed_upload_file(file) as path:
        return file_download(repair_pdf(path), "application/pdf", "repaired.pdf")



@app.post("/api/convert/image-to-pdf")
async def api_image_to_pdf(request: Request):
    form = await request.form()
    selected_files = [
        upload
        for field_name in ("file", "files")
        for upload in form.getlist(field_name)
        if hasattr(upload, "filename") and hasattr(upload, "file")
    ]
    return file_download(image_to_pdf(selected_files), "application/pdf", "converted.pdf")


@app.post("/api/convert/html-to-pdf")
async def api_html_to_pdf(file: UploadFile = File(...)):
    with managed_upload_file(file) as path:
        return file_download(html_to_pdf(path), "application/pdf", "html-converted.pdf")


@app.post("/api/convert/word-to-pdf")
async def api_word_to_pdf(file: UploadFile = File(...)):
    with managed_upload_file(file) as path:
        return file_download(word_to_pdf(path), "application/pdf", "word-converted.pdf")


@app.post("/api/convert/excel-to-pdf")
async def api_excel_to_pdf(file: UploadFile = File(...)):
    with managed_upload_file(file) as path:
        return file_download(excel_to_pdf(path), "application/pdf", "excel-converted.pdf")


@app.post("/api/convert/powerpoint-to-pdf")
async def api_powerpoint_to_pdf(file: UploadFile = File(...)):
    with managed_upload_file(file) as path:
        return file_download(powerpoint_to_pdf(path), "application/pdf", "powerpoint-converted.pdf")


@app.post("/api/delete")
async def api_delete(file: UploadFile = File(...), pages: str = Form(...)):
    with managed_upload_file(file) as path:
        return file_download(delete_pdf_pages(path, pages), "application/pdf", "deleted.pdf")



@app.post("/api/extract")
async def api_extract(file: UploadFile = File(...), pages: str = Form(...)):
    with managed_upload_file(file) as path:
        return file_download(extract_pdf_pages(path, pages), "application/pdf", "extracted.pdf")


@app.post("/api/ocr")
async def api_ocr(file: UploadFile = File(...)):
    with managed_upload_file(file) as path:
        return file_download(extract_text_from_pdf(path), "text/plain; charset=utf-8", "ocr.txt")


@app.post("/api/watermark")
async def api_watermark(file: UploadFile = File(...), text: str = Form(...), opacity: float = Form(0.3)):
    # P1.3: Validate opacity bounds
    if not (0 < opacity <= 1):
        raise HTTPException(status_code=400, detail="opacity must be between 0 (exclusive) and 1 (inclusive)")
    with managed_upload_file(file) as path:
        return file_download(add_watermark(path, text, opacity), "application/pdf", "watermarked.pdf")


@app.post("/api/pdf-to-jpg")
async def api_pdf_to_jpg(file: UploadFile = File(...)):
    with managed_upload_file(file) as path:
        return file_download(pdf_to_images(path), "application/zip", "images.zip")


@app.post("/api/pdf-to-word")
async def api_pdf_to_word(file: UploadFile = File(...)):
    with managed_upload_file(file) as path:
        return file_download(
            pdf_to_word(path),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "converted.docx",
        )


@app.post("/api/pdf-to-excel")
async def api_pdf_to_excel(file: UploadFile = File(...)):
    with managed_upload_file(file) as path:
        return file_download(
            pdf_to_excel(path),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "converted.xlsx",
        )


@app.post("/api/pdf-to-powerpoint")
async def api_pdf_to_powerpoint(file: UploadFile = File(...)):
    with managed_upload_file(file) as path:
        return file_download(
            pdf_to_powerpoint(path),
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            "converted.pptx",
        )



@app.post("/api/numbering")
async def api_numbering(
    file: UploadFile = File(...),
    format_str: str = Form("{page}"),
    position: str = Form("bottom-right"),
):
    with managed_upload_file(file) as path:
        return file_download(add_page_numbers(path, format_str, position), "application/pdf", "numbered.pdf")


@app.post("/api/protect")
async def api_protect(
    file: UploadFile = File(...),
    user_password: str = Form(...),
    owner_password: Optional[str] = Form(None),
):
    with managed_upload_file(file) as path:
        return file_download(protect_pdf(path, user_password, owner_password), "application/pdf", "protected.pdf")


@app.post("/api/unlock")
async def api_unlock(file: UploadFile = File(...), password: str = Form(...)):
    with managed_upload_file(file) as path:
        return file_download(unlock_pdf(path, password), "application/pdf", "unlocked.pdf")



@app.post("/api/compare")
async def api_compare(file_a: UploadFile = File(...), file_b: UploadFile = File(...)):
    with managed_upload_file(file_a) as path_a:
        with managed_upload_file(file_b) as path_b:
            return file_download(compare_pdfs(path_a, path_b), "application/json", "compare-report.json")


@app.post("/api/censor")
async def api_censor(
    file: UploadFile = File(...),
    terms: str = Form(...),
    case_sensitive: bool = Form(False),
):
    with managed_upload_file(file) as path:
        return file_download(censor_pdf(path, terms, case_sensitive), "application/pdf", "censored.pdf")


@app.post("/api/sign")
async def api_sign(
    file: UploadFile = File(...),
    signer_name: str = Form(...),
    reason: str = Form(""),
    location: str = Form(""),
    position: str = Form("bottom-right"),
):
    with managed_upload_file(file) as path:
        return file_download(sign_pdf(path, signer_name, reason, location, position), "application/pdf", "signed.pdf")


@app.post("/api/ai/summarize")
async def api_summarize(file: UploadFile = File(...), max_sentences: int = Form(6)):
    with managed_upload_file(file) as path:
        return file_download(summarize_pdf(path, max_sentences), "text/plain; charset=utf-8", "summary.txt")


@app.post("/api/ai/translate")
async def api_translate(file: UploadFile = File(...), target_language: str = Form("francais")):
    with managed_upload_file(file) as path:
        return file_download(translate_pdf(path, target_language), "application/pdf", "translated.pdf")

