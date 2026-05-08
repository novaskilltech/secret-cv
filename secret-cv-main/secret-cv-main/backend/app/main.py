from fastapi import FastAPI, Depends, File, UploadFile, HTTPException, status
from fastapi.responses import StreamingResponse
from typing import List
import io

from .anonymizer import anonymize_pdf
from .schemas import AnonymizeResponse
from .dependencies import get_current_user

app = FastAPI(
    title="CV Anonymizer API",
    description="Endpoint to anonymize personal data in PDF CVs.",
    version="1.0.0",
)

@app.post("/api/anonymize", response_model=AnonymizeResponse, status_code=status.HTTP_200_OK)
async def anonymize_endpoint(
    file: UploadFile = File(...),
    user=Depends(get_current_user),
):
    """Receive a PDF, anonymize it, and return the processed file.
    The actual anonymization logic is delegated to ``anonymize_pdf``.
    """
    if file.content_type != "application/pdf":
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Only PDF files are supported.",
        )
    raw_bytes = await file.read()
    try:
        anonymized_bytes = anonymize_pdf(raw_bytes)
    except Exception as exc:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail=str(exc),
        )
    # Return the anonymized PDF as streaming response
    return StreamingResponse(
        io.BytesIO(anonymized_bytes),
        media_type="application/pdf",
        headers={"Content-Disposition": f"attachment; filename=anonymized_{file.filename}"},
    )
