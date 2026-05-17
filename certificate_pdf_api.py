from io import BytesIO
from pathlib import Path
from uuid import uuid4
from urllib.parse import quote

from fastapi import APIRouter, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from pypdf import PdfReader, PdfWriter


router = APIRouter()

CERT_TEMPLATE_FILE = Path("certificate_template.pdf")
GENERATED_DIR = Path("generated_files")
GENERATED_DIR.mkdir(exist_ok=True)


class ExportCertificatePdfRequest(BaseModel):
    filename: str
    prepared_for: str
    property_address: str
    market_estimate_low: str
    market_estimate_high: str
    recommended_launch_price: str
    property_practitioner: str
    broker_owner_manager: str
    certificate_date: str
    office_name: str


def draw_centered(c, text, x, y, font_size=10):
    c.setFont("Helvetica", font_size)
    c.drawCentredString(x, y, text)


def build_certificate_pdf(payload: ExportCertificatePdfRequest) -> tuple[bytes, str]:
    if not CERT_TEMPLATE_FILE.exists():
        raise HTTPException(status_code=500, detail="certificate_template.pdf not found on server.")

    packet = BytesIO()
    c = canvas.Canvas(packet, pagesize=A4)

    # First-pass coordinates. We will adjust after the next test.
    draw_centered(c, payload.prepared_for, 321, 616, 10)
    draw_centered(c, payload.property_address, 321, 561, 10)

    market_range = f"{payload.market_estimate_low} to {payload.market_estimate_high}"
    draw_centered(c, market_range, 350, 456, 12)

    draw_centered(c, payload.recommended_launch_price, 350, 346, 14)

    draw_centered(c, payload.property_practitioner, 205, 205, 10)
    draw_centered(c, payload.broker_owner_manager, 445, 205, 10)

    draw_centered(c, payload.certificate_date, 205, 130, 10)
    draw_centered(c, payload.office_name, 445, 130, 10)

    c.save()
    packet.seek(0)

    template_pdf = PdfReader(str(CERT_TEMPLATE_FILE))
    overlay_pdf = PdfReader(packet)

    writer = PdfWriter()
    page = template_pdf.pages[0]
    page.merge_page(overlay_pdf.pages[0])
    writer.add_page(page)

    output = BytesIO()
    writer.write(output)

    safe_filename = payload.filename if payload.filename.endswith(".pdf") else f"{payload.filename}.pdf"
    return output.getvalue(), safe_filename


@router.post("/export/certificate-pdf/gpt-link")
def export_certificate_pdf_gpt_link(payload: ExportCertificatePdfRequest):
    file_bytes, safe_filename = build_certificate_pdf(payload)

    unique_name = f"{uuid4()}_{safe_filename}"
    file_path = GENERATED_DIR / unique_name
    file_path.write_bytes(file_bytes)

    encoded_name = quote(unique_name)
    download_url = f"https://valuation-export-api.onrender.com/download-certificate-pdf/{encoded_name}"

    return {
        "openaiFileResponse": [
            {
                "name": safe_filename,
                "mime_type": "application/pdf",
                "download_link": download_url
            }
        ],
        "filename": safe_filename,
        "status": "success"
    }


@router.get("/download-certificate-pdf/{file_name}")
def download_certificate_pdf(file_name: str):
    file_path = GENERATED_DIR / file_name

    if not file_path.exists():
        raise HTTPException(status_code=404, detail="File not found.")

    original_name = file_name.split("_", 1)[1] if "_" in file_name else file_name

    return FileResponse(
        path=file_path,
        media_type="application/pdf",
        filename=original_name
    )