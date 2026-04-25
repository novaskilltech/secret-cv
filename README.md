# NOVA PDF Tools

Prototype open-source d'un editeur PDF web inspire de ilovepdf.com.

## Stack
- Backend: FastAPI
- Frontend: HTML/CSS/JavaScript statique
- PDF processing: PyPDF2, pikepdf, reportlab, pdfplumber
- PDF vers images: pdf2image avec fallback pypdfium2
- Conversions bureautiques: python-docx, openpyxl, python-pptx

## Fonctionnalites livrees

### Organiser PDF
- Fusionner PDF: `POST /api/merge`
- Diviser PDF: `POST /api/split`
- Supprimer des pages: `POST /api/delete`
- Extraire des pages: `POST /api/extract`
- Reordonner les pages: `POST /api/reorder`

### Optimiser le PDF
- Compresser PDF: `POST /api/compress`
- Reparer PDF: `POST /api/repair`
- OCR / extraction texte native: `POST /api/ocr`

### Convertir en PDF
- Image vers PDF: `POST /api/convert/image-to-pdf`
- HTML vers PDF: `POST /api/convert/html-to-pdf`
- Word vers PDF: `POST /api/convert/word-to-pdf`
- Excel vers PDF: `POST /api/convert/excel-to-pdf`
- PowerPoint vers PDF: `POST /api/convert/powerpoint-to-pdf`

### Convertir depuis PDF
- PDF vers JPG: `POST /api/pdf-to-jpg`
- PDF vers Word: `POST /api/pdf-to-word`
- PDF vers Excel: `POST /api/pdf-to-excel`
- PDF vers PowerPoint: `POST /api/pdf-to-powerpoint`

### Modifier PDF
- Faire pivoter PDF: `POST /api/rotate`
- Ajouter des numeros de pages: `POST /api/numbering`
- Ajouter un filigrane: `POST /api/watermark`
- Rogner PDF: `POST /api/crop`

### Securite PDF
- Deverrouiller PDF: `POST /api/unlock`
- Proteger PDF: `POST /api/protect`
- Comparer PDF: `POST /api/compare`
- Censurer PDF: `POST /api/censor`

## Notes importantes
- Les conversions `Word/Excel/PowerPoint/HTML -> PDF` utilisent automatiquement `LibreOffice` quand il est disponible sur le serveur ou dans le conteneur.
- Si `LibreOffice` n'est pas disponible ou echoue sur un fichier, l'application retombe sur un rendu simplifie base sur le texte et les tableaux extraits.
- `PDF -> Word/Excel/PowerPoint` reste un export structure simplifie:
  - Word: texte extrait page par page
  - Excel: tables detectees ou texte par page
  - PowerPoint: une image par page PDF
- `OCR PDF` utilise `OCRmyPDF` quand il est disponible pour traiter les PDF scannes ou images.
- Si `OCRmyPDF` n'est pas disponible ou echoue, la route retombe sur l'extraction de texte native du PDF.
- `Censurer PDF` est maintenant disponible en mode irreversible par rasterisation des pages touchees et reconstruction du PDF.
- Signature numerique, resume IA et traduction IA ne sont pas encore livres.

## Installation
```bash
cd "c:/Users/P C/Downloads/PDF TOOLS/backend"
pip install -r requirements.txt
uvicorn main:app --reload --host 0.0.0.0 --port 8000
```

Ouvrir ensuite `http://localhost:8000`.

## Docker
Le projet inclut maintenant un environnement conteneurise avec moteurs externes cote serveur:
- LibreOffice
- Tesseract OCR
- OCRmyPDF
- Ghostscript
- Poppler

Lancement:
```bash
cd "c:/Users/P C/Downloads/PDF TOOLS"
docker compose up --build
```

Arret:
```bash
docker compose down
```

L'application sera disponible sur `http://localhost:8000`.

Fichiers lies:
- `docker-compose.yml`
- `backend/Dockerfile`
- `.dockerignore`

Notes:
- les moteurs externes sont installes dans le conteneur, pas chez les utilisateurs
- `tmpfs` est utilise pour les fichiers temporaires du traitement PDF
- le backend appelle deja automatiquement `LibreOffice` pour les conversions vers PDF et `OCRmyPDF` pour l'OCR quand ces moteurs sont disponibles

## Tests
```bash
cd "c:/Users/P C/Downloads/PDF TOOLS/backend"
python test_endpoints.py
```

## Verification actuelle
- Validation serveur des plages de pages et parametres critiques
- Limite serveur par upload via `NOVA_MAX_UPLOAD_BYTES` (50 Mo par defaut)
- Timeout des commandes externes via `NOVA_COMMAND_TIMEOUT_SECONDS` (120 s par defaut)
- Origines CORS configurables via `NOVA_CORS_ORIGINS` (localhost uniquement par defaut)
- Conteneur Docker execute avec un utilisateur non-root
- Drag and drop sur le frontend
- Telechargement automatique du resultat
- Tests d'integration backend: `21/21` verts

## Roadmap restante

### Blocages techniques a lever
- Signature PDF avec certificats
- OCR image complet
- Resume et traduction par IA
- Conversions Office haute fidelite via moteur externe
