from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import tempfile
from pathlib import Path

pdf_path = Path(tempfile.gettempdir()) / "test_input.pdf"
c = canvas.Canvas(str(pdf_path), pagesize=letter)
for page_num in range(1, 4):
    c.drawString(100, 750, f"Test Page {page_num}")
    c.drawString(100, 700, "Test PDF for endpoints")
    c.showPage()
c.save()
print(f"✓ Test PDF: {pdf_path} ({pdf_path.stat().st_size} bytes)")
