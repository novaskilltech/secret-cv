FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    DEBIAN_FRONTEND=noninteractive \
    LIBREOFFICE_HEADLESS=1 \
    TESSDATA_PREFIX=/usr/share/tesseract-ocr/5/tessdata

WORKDIR /app

RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    libreoffice-writer \
    libreoffice-calc \
    libreoffice-impress \
    libreoffice-java-common \
    poppler-utils \
    tesseract-ocr \
    tesseract-ocr-eng \
    tesseract-ocr-fra \
    ghostscript \
    ocrmypdf \
    unpaper \
    pngquant \
    qpdf \
    fonts-dejavu \
    fonts-liberation \
    fonts-noto-core \
    ca-certificates \
    curl \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt ./requirements.txt
RUN python -m pip install --upgrade pip && pip install -r requirements.txt

COPY . .
RUN chmod +x entrypoint.sh && \
    useradd --create-home --shell /usr/sbin/nologin appuser && \
    chown -R appuser:appuser /app

USER appuser

# EXPOSE and HEALTHCHECK are managed by Cloud Run
ENTRYPOINT ["./entrypoint.sh"]
# Start with dynamic port for Cloud Run compatibility
CMD uvicorn main:app --host 0.0.0.0 --port ${PORT:-8000}
