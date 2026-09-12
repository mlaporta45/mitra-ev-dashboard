import io
import hashlib
from pathlib import Path

MAX_CHARS = 150_000


def file_hash(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()[:16]


def extract_from_pdf(data: bytes) -> str:
    import pdfplumber
    parts = []
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            if text.strip():
                parts.append(f"[Page {i+1}]\n{text}")
            for table in page.extract_tables():
                rows = []
                for row in table:
                    rows.append(" | ".join(str(c) if c else "" for c in row))
                if rows:
                    parts.append("[Table]\n" + "\n".join(rows))
    return "\n\n".join(parts)


def extract_from_docx(data: bytes) -> str:
    from docx import Document
    doc = Document(io.BytesIO(data))
    parts = []
    for para in doc.paragraphs:
        if para.text.strip():
            parts.append(para.text)
    for table in doc.tables:
        rows = []
        for row in table.rows:
            rows.append(" | ".join(c.text.strip() for c in row.cells))
        parts.append("[Table]\n" + "\n".join(rows))
    return "\n".join(parts)


def extract_from_pptx(data: bytes) -> str:
    from pptx import Presentation
    prs = Presentation(io.BytesIO(data))
    parts = []
    for i, slide in enumerate(prs.slides):
        slide_parts = [f"[Slide {i+1}]"]
        for shape in slide.shapes:
            if shape.has_text_frame:
                text = "\n".join(p.text for p in shape.text_frame.paragraphs if p.text.strip())
                if text:
                    slide_parts.append(text)
            if shape.has_table:
                rows = []
                for row in shape.table.rows:
                    rows.append(" | ".join(c.text.strip() for c in row.cells))
                slide_parts.append("[Table]\n" + "\n".join(rows))
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text.strip()
            if notes:
                slide_parts.append(f"[Speaker Notes]\n{notes}")
        parts.append("\n".join(slide_parts))
    return "\n\n".join(parts)


def extract_from_excel(data: bytes) -> str:
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    parts = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = []
        for row in ws.iter_rows(values_only=True):
            if any(c is not None for c in row):
                rows.append(" | ".join(str(c) if c is not None else "" for c in row))
        if rows:
            parts.append(f"[Sheet: {sheet_name}]\n" + "\n".join(rows))
    return "\n\n".join(parts)


def extract_text(filename: str, data: bytes) -> str:
    ext = Path(filename).suffix.lower()
    if ext == ".pdf":
        text = extract_from_pdf(data)
    elif ext in (".docx", ".doc"):
        text = extract_from_docx(data)
    elif ext in (".pptx", ".ppt"):
        text = extract_from_pptx(data)
    elif ext in (".xlsx", ".xlsm", ".xls", ".xlsb"):
        text = extract_from_excel(data)
    else:
        text = data.decode("utf-8", errors="replace")
    return text[:MAX_CHARS]
