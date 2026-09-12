import io
import re
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


_SKIP_SHEET_KEYWORDS = ("gl tb", "general ledger", "trial balance", "gl detail", "raw data")


def _precompute_mitra_metrics(wb) -> tuple[str, list[str]]:
    """
    Pre-compute metrics that require counting/averaging across rows.
    Returns (pre-computed summary block, list of diagnostic messages).
    """
    computed = []
    diagnostics = []
    all_sheets = wb.sheetnames
    diagnostics.append(f"Excel sheets found: {all_sheets}")

    # --- Vehicles in Service: Fleet List sheet, column M = "VEH. STATUS", count "In-Service" ---
    fleet_sheet = next(
        (s for s in wb.sheetnames if "fleet list" in s.lower() or re.search(r'\bpart\s+i\b', s, re.IGNORECASE)),
        None,
    )
    diagnostics.append(f"Fleet sheet match: {fleet_sheet!r}")
    if fleet_sheet:
        ws = wb[fleet_sheet]
        header_row = None
        status_col = None
        for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
            if row and any("status" in str(c).lower() for c in row if c):
                header_row = i
                for j, c in enumerate(row):
                    if c and "veh. status" in str(c).lower():
                        status_col = j
                        break
                if status_col is not None:
                    break
        if status_col is not None and header_row is not None:
            in_service = 0
            for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
                val = row[status_col] if len(row) > status_col else None
                if val and "in-service" in str(val).lower():
                    in_service += 1
            computed.append(f"PRE-COMPUTED Vehicles in Service (counted from Fleet List, VEH. STATUS = In-Service): {in_service}")

    # --- DCFC In Service + Avg Utilization: Part IV DCFC Utilization sheet ---
    dcfc_sheet = next(
        (s for s in wb.sheetnames if "dcfc" in s.lower() and "util" in s.lower()),
        None,
    )
    if not dcfc_sheet:
        # Broaden search: any sheet with "dcfc"
        dcfc_sheet = next((s for s in wb.sheetnames if "dcfc" in s.lower()), None)
    diagnostics.append(f"DCFC sheet match: {dcfc_sheet!r}")
    if dcfc_sheet:
        ws = wb[dcfc_sheet]
        # Find header row (has "startTimestamp" and "Utilization")
        header_row = None
        ts_col = util_col = depot_col = None
        for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
            if row and any(str(c).lower() == "starttimestamp" for c in row if c):
                header_row = i
                for j, c in enumerate(row):
                    if c:
                        cl = str(c).lower()
                        if cl == "starttimestamp":
                            ts_col = j
                        elif cl == "utilization":
                            util_col = j
                        elif "depot" in cl or "id" in cl:
                            depot_col = j
                break
        if header_row and ts_col is not None and util_col is not None:
            from collections import defaultdict
            month_data = defaultdict(list)
            for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
                if not row or row[ts_col] is None:
                    continue
                ts = str(row[ts_col]).strip()
                util = row[util_col]
                if ts and util is not None:
                    try:
                        month_data[ts].append(float(util))
                    except (ValueError, TypeError):
                        pass
            if month_data:
                latest_month = sorted(month_data.keys())[-1]
                utils = month_data[latest_month]
                dcfc_count = len(utils)
                avg_util = sum(utils) / len(utils) if utils else 0
                computed.append(f"PRE-COMPUTED DCFC In Service (count of charger entries for latest month {latest_month}): {dcfc_count}")
                computed.append(f"PRE-COMPUTED DCFC Avg. Utilization Rate for {latest_month}: {avg_util:.4f} (as decimal, e.g. 0.0032 = 0.32%)")

    # --- EBITDA: Income Statement sheet — Net Income + Depreciation + Amortization + Interest + Tax ---
    is_sheet = next(
        (s for s in wb.sheetnames if "income statement" in s.lower()),
        None,
    )
    diagnostics.append(f"Income Statement sheet match: {is_sheet!r}")
    if is_sheet:
        ws = wb[is_sheet]
        # Find the most recent period column (last non-empty column in header row)
        header_row_idx = None
        months = ("jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec")
        for i, row in enumerate(ws.iter_rows(min_row=1, max_row=15, values_only=True), start=1):
            if not row:
                continue
            # Require at least 3 cells containing month names (confirms it's the column header row)
            hits = sum(1 for c in row if c and any(m in str(c).lower() for m in months))
            if hits >= 3:
                header_row_idx = i
                break
        if header_row_idx:
            header = list(ws.iter_rows(min_row=header_row_idx, max_row=header_row_idx, values_only=True))[0]
            # Find rightmost period column — skip "Total", booleans, and empty cells
            last_col = None
            for j in range(len(header) - 1, 1, -1):
                val = header[j]
                if val is None or isinstance(val, bool):
                    continue
                sv = str(val).lower()
                if "total" in sv or "financial row" in sv:
                    continue
                last_col = j
                break
            period_label = str(header[last_col]) if last_col else "most recent"

            ebitda_components = {}
            keywords = {
                "net income": "net_income",
                "net loss": "net_income",
                "depreciation": "depreciation",
                "amortization": "amortization",
                "interest": "interest",
                "taxes & licenses": "tax",
                "tax": "tax",
            }
            for row in ws.iter_rows(values_only=True):
                label = str(row[1]).lower() if len(row) > 1 and row[1] else ""
                for kw, key in keywords.items():
                    if kw in label and key not in ebitda_components:
                        val = row[last_col] if last_col and len(row) > last_col else None
                        if val is not None:
                            try:
                                ebitda_components[key] = float(val)
                            except (ValueError, TypeError):
                                pass
            if "net_income" in ebitda_components:
                ebitda = sum(ebitda_components.get(k, 0) for k in ["net_income", "depreciation", "amortization", "interest", "tax"])
                detail = " + ".join(f"{k}={v:.2f}" for k, v in ebitda_components.items())
                computed.append(f"PRE-COMPUTED EBITDA for {period_label} = {ebitda:.2f} (calculated as: {detail})")

    block = ""
    if computed:
        block = "[PRE-COMPUTED METRICS — use these values directly, do not recalculate]\n" + "\n".join(computed) + "\n"
    return block, diagnostics


_last_excel_diagnostics: list[str] = []


def extract_from_excel(data: bytes) -> str:
    global _last_excel_diagnostics
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    precomputed, diagnostics = _precompute_mitra_metrics(wb)
    _last_excel_diagnostics = diagnostics
    parts = []
    for sheet_name in wb.sheetnames:
        # Skip raw general ledger / trial balance sheets — too large, not useful for KPI extraction
        if any(kw in sheet_name.lower() for kw in _SKIP_SHEET_KEYWORDS):
            continue
        ws = wb[sheet_name]
        rows = []
        for row in ws.iter_rows(values_only=True):
            if any(c is not None for c in row):
                rows.append(" | ".join(str(c) if c is not None else "" for c in row))
        if rows:
            parts.append(f"[Sheet: {sheet_name}]\n" + "\n".join(rows))
    full_text = "\n\n".join(parts)
    combined = (precomputed + "\n\n" + full_text) if precomputed else full_text
    return combined[:MAX_CHARS]


def get_last_excel_diagnostics() -> list[str]:
    return list(_last_excel_diagnostics)


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
