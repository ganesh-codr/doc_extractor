import os
import re
from pathlib import Path

from openpyxl import Workbook
from pypdf import PdfReader

try:
    import pytesseract
    from pdf2image import convert_from_path
    from pdf2image.exceptions import PDFInfoNotInstalledError
except Exception:  # pragma: no cover - optional OCR dependencies
    pytesseract = None
    convert_from_path = None
    PDFInfoNotInstalledError = Exception


def extract_text(source, source_type="pdf"):
    """Return plain text from a PDF path, a scanned PDF path, or a raw text string."""
    if source_type in ("text", "txt"):
        path = Path(source)
        if path.exists():
            return path.read_text(encoding="utf-8", errors="ignore")
        return str(source)

    path = Path(source)
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")

    text_chunks = []
    try:
        reader = PdfReader(str(path))
        for page in reader.pages:
            page_text = page.extract_text() or ""
            text_chunks.append(page_text)
    except Exception:
        text_chunks = []

    text = "\n".join(chunk.strip() for chunk in text_chunks if chunk and chunk.strip())

    if text.strip():
        return text

    if pytesseract is None or convert_from_path is None:
        raise RuntimeError("OCR dependencies are not available. Install pytesseract and pdf2image.")

    try:
        images = convert_from_path(str(path))
    except PDFInfoNotInstalledError as exc:
        raise RuntimeError(
            "Scanned PDF OCR needs Poppler installed and available on PATH. "
            "Install Poppler for Windows and retry the extraction."
        ) from exc

    for image in images:
        text_chunks.append(pytesseract.image_to_string(image))

    return "\n".join(chunk.strip() for chunk in text_chunks if chunk and chunk.strip())


def build_schedule_rows(text, prompt):
    """Build a simple cable-schedule style table from the drawing text."""
    lines = [line.strip() for line in text.splitlines() if line.strip()]

    rows = []
    for line in lines:
        lowered = line.lower()
        if "cable" not in lowered and "tag" not in lowered and "from" not in lowered and "to" not in lowered:
            continue

        tag_match = re.search(r"(?:cable\s*tag|tag)\s+([A-Za-z0-9_-]+)", line, re.I)
        cable_tag = tag_match.group(1) if tag_match else ""

        from_match = re.search(r"from\s+([A-Za-z0-9/-]+)", line, re.I)
        to_match = re.search(r"to\s+([A-Za-z0-9/-]+)", line, re.I)
        size_match = re.search(r"size\s+([A-Za-z0-9\-./]+)", line, re.I)
        type_match = re.search(r"type\s+([A-Za-z0-9\-./]+)", line, re.I)

        from_equipment = from_match.group(1) if from_match else ""
        to_equipment = to_match.group(1) if to_match else ""

        rows.append({
            "Sl.no": len(rows) + 1,
            "Cable tag": cable_tag or "N/A",
            "From Equipment": from_equipment,
            "From Location": "",
            "To Equipment": to_equipment,
            "To Location": "",
            "Cable Size": size_match.group(1) if size_match else "",
            "Cable Type": type_match.group(1) if type_match else "",
            "Remarks": "Extracted from prompt: " + prompt[:120],
        })

    if not rows:
        rows.append({
            "Sl.no": 1,
            "Cable tag": "N/A",
            "From Equipment": "",
            "From Location": "",
            "To Equipment": "",
            "To Location": "",
            "Cable Size": "",
            "Cable Type": "",
            "Remarks": "No cable schedule data found. Please refine the prompt.",
        })

    return rows


def paginate_rows(rows, page_number=1, page_size=25):
    """Return a single page of rows with pagination metadata."""
    total_rows = len(rows)
    page_count = max(1, (total_rows + page_size - 1) // page_size)
    page_number = max(1, min(int(page_number), page_count))
    start = (page_number - 1) * page_size
    end = min(start + page_size, total_rows)

    return {
        "page": page_number,
        "page_count": page_count,
        "page_size": page_size,
        "total_rows": total_rows,
        "rows": rows[start:end],
    }


def export_excel(rows, output_path):
    """Write extracted rows to an Excel workbook."""
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Extraction"

    headers = [
        "Sl.no",
        "Cable tag",
        "From Equipment",
        "From Location",
        "To Equipment",
        "To Location",
        "Cable Size",
        "Cable Type",
        "Remarks",
    ]
    sheet.append(headers)

    for row in rows:
        sheet.append([row.get(column, "") for column in headers])

    workbook.save(output)
    return str(output)


def export_pdf(rows, output_path):
    """Write a simple PDF report of the extracted rows."""
    try:
        from fpdf import FPDF
    except Exception as exc:  # pragma: no cover - optional dependency
        raise RuntimeError("PDF export requires the 'fpdf' package.") from exc

    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "Engineering Drawing Extraction", ln=True)
    pdf.set_font("Arial", size=9)

    for row in rows:
        line = (
            f"{row.get('Sl.no', '')} | {row.get('Cable tag', '')} | {row.get('From Equipment', '')} | "
            f"{row.get('To Equipment', '')} | {row.get('Cable Size', '')} | {row.get('Cable Type', '')}"
        )
        pdf.multi_cell(0, 6, line)

    pdf.output(str(output))
    return str(output)
