"""
DocuSense AI — excel_builder.py

Builds the 3-sheet Excel report (Summary Dashboard, Header Details, Line
Items) from extracted documents, using openpyxl with the house style:
dark-blue bold headers, alternating row shading, Arial 10pt throughout.

Every sheet-building function tolerates empty input (no documents) so a
partial run — where some files failed — still produces a complete,
downloadable workbook.
"""

from __future__ import annotations

from datetime import datetime
from io import BytesIO
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from analyzer import normalize_amount

FONT_NAME = "Arial"
FONT_SIZE = 10

HEADER_FILL = PatternFill(start_color="FF1F4E79", end_color="FF1F4E79", fill_type="solid")
HEADER_FONT = Font(name=FONT_NAME, size=FONT_SIZE, bold=True, color="FFFFFFFF")
ALT_FILL = PatternFill(start_color="FFDEEAF1", end_color="FFDEEAF1", fill_type="solid")
WHITE_FILL = PatternFill(start_color="FFFFFFFF", end_color="FFFFFFFF", fill_type="solid")
BODY_FONT = Font(name=FONT_NAME, size=FONT_SIZE)
TITLE_FONT = Font(name=FONT_NAME, size=14, bold=True, color="FF1F4E79")
SECTION_FONT = Font(name=FONT_NAME, size=11, bold=True, color="FF1F4E79")

WRAP_TOP_LEFT = Alignment(wrap_text=True, vertical="top")


# ---------------------------------------------------------------------------
# Generic styling helpers
# ---------------------------------------------------------------------------

def _write_header_row(ws: Worksheet, row: int, headers: list[str], start_col: int = 1) -> None:
    for i, header in enumerate(headers):
        cell = ws.cell(row=row, column=start_col + i, value=header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical="center")


def _to_cell_value(value: Any) -> Any:
    """openpyxl can only write str/int/float/bool/datetime/None — flatten anything else (e.g. list-type custom fields)."""
    if isinstance(value, (list, tuple, set)):
        return ", ".join(str(v) for v in value)
    if isinstance(value, dict):
        return str(value)
    return value


def _write_data_row(ws: Worksheet, row: int, values: list[Any], alt: bool, start_col: int = 1) -> None:
    fill = ALT_FILL if alt else WHITE_FILL
    for i, value in enumerate(values):
        cell = ws.cell(row=row, column=start_col + i, value=_to_cell_value(value))
        cell.fill = fill
        cell.font = BODY_FONT
        cell.alignment = WRAP_TOP_LEFT


def _write_table(ws: Worksheet, start_row: int, headers: list[str], rows: list[list[Any]], start_col: int = 1) -> int:
    """Writes a header row + alternating data rows starting at start_row. Returns the next free row."""
    _write_header_row(ws, start_row, headers, start_col)
    row = start_row + 1
    for i, values in enumerate(rows):
        _write_data_row(ws, row, values, alt=(i % 2 == 0), start_col=start_col)
        row += 1
    if not rows:
        row += 1  # leave a visible empty row under the header so empty tables aren't confusing
    return row


def _section_title(ws: Worksheet, row: int, text: str) -> int:
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = SECTION_FONT
    return row + 1


def _autosize_columns(ws: Worksheet, max_col: int, min_width: int = 12, max_width: int = 55) -> None:
    for col_idx in range(1, max_col + 1):
        longest = min_width
        for cell in ws[get_column_letter(col_idx)]:
            if cell.value is None:
                continue
            longest = max(longest, min(len(str(cell.value)), max_width))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(longest + 2, max_width)


def _ok_docs(documents: list[dict]) -> list[dict]:
    return [d for d in documents if d.get("success") and isinstance(d.get("data"), dict)]


# ---------------------------------------------------------------------------
# Sheet 1 — Summary Dashboard
# ---------------------------------------------------------------------------

def _build_summary_dashboard(wb: Workbook, documents: list[dict], generated_at: datetime) -> None:
    ws = wb.create_sheet("Summary Dashboard")
    ok_docs = _ok_docs(documents)
    failed_docs = [d for d in documents if not d.get("success")]

    ws.cell(row=1, column=1, value="DocuSense AI — Summary Dashboard").font = TITLE_FONT
    ws.cell(row=2, column=1, value=f"Processed: {generated_at.strftime('%Y-%m-%d %H:%M:%S')}").font = BODY_FONT

    row = 4
    row = _section_title(ws, row, "Documents by Type")
    by_type: dict[str, dict[str, float]] = {}
    for d in ok_docs:
        amount = normalize_amount(d["data"].get("total")) or 0.0
        entry = by_type.setdefault(d["doc_type"], {"count": 0, "total": 0.0})
        entry["count"] += 1
        entry["total"] += amount

    type_rows = [[doc_type, entry["count"], round(entry["total"], 2)] for doc_type, entry in sorted(by_type.items())]
    row = _write_table(ws, row, ["Document Type", "Count", "Total Value"], type_rows)
    row += 1

    row = _section_title(ws, row, "Overall Stats")
    stat_rows = [
        ["Total Files Uploaded", len(documents)],
        ["Successful Extractions", len(ok_docs)],
        ["Failed Extractions", len(failed_docs)],
        ["Processing Date", generated_at.strftime("%Y-%m-%d")],
        ["Processing Time", generated_at.strftime("%H:%M:%S")],
    ]
    row = _write_table(ws, row, ["Metric", "Value"], stat_rows)

    _autosize_columns(ws, 5)


# ---------------------------------------------------------------------------
# Sheet 2 — Header Details
# ---------------------------------------------------------------------------

# (Excel column header, source JSON key) — one row per document.
HEADER_DETAIL_COLUMNS = [
    ("Document Type", "document_type"),
    ("Invoice Number", "invoice_number"),
    ("Invoice Date", "invoice_date"),
    ("PO Number", "po_number"),
    ("Vendor Name", "vendor_name"),
    ("Bill To", "bill_to"),
    ("Subtotal", "subtotal"),
    ("Tax", "tax"),
    ("Shipping Charges", "shipping_charges"),
    ("Total", "total"),
    ("Payment Terms", "payment_terms"),
    ("Due Date", "due_date"),
    ("Bank Details", "bank_details"),
]


def header_detail_row(d: dict, custom_field_names: list[str] | None = None) -> dict:
    """One Header Details row as a {column_label: value} dict — shared by the Excel sheet and uipath_bridge.py."""
    custom_field_names = custom_field_names or []
    data = d.get("data") if d.get("success") and isinstance(d.get("data"), dict) else None
    row = {"File Name": d["filename"]}
    for label, key in HEADER_DETAIL_COLUMNS:
        row[label] = data.get(key) if data else None
    for name in custom_field_names:
        row[name] = data.get(name) if data else None
    return row


def _build_header_details_sheet(wb: Workbook, documents: list[dict], custom_field_names: list[str] | None = None) -> None:
    custom_field_names = custom_field_names or []
    ws = wb.create_sheet("Header Details")
    headers = ["File Name"] + [label for label, _ in HEADER_DETAIL_COLUMNS] + custom_field_names
    rows = [[header_detail_row(d, custom_field_names)[h] for h in headers] for d in documents]

    _write_table(ws, 1, headers, rows)
    ws.freeze_panes = "A2"
    _autosize_columns(ws, len(headers))


# ---------------------------------------------------------------------------
# Sheet 3 — Line Items
# ---------------------------------------------------------------------------

LINE_ITEM_COLUMNS = ["Source Doc", "Doc Number", "Line#", "Part Number", "Description", "Qty", "Unit Price", "Tax", "Line Amount"]


def line_item_rows(d: dict) -> list[dict]:
    """This document's Line Items rows as {column_label: value} dicts — shared by the Excel sheet and uipath_bridge.py."""
    if not (d.get("success") and isinstance(d.get("data"), dict)):
        return []
    data = d["data"]
    line_items = data.get("line_items")
    if not isinstance(line_items, list):
        return []

    doc_number = data.get("invoice_number") or ""
    rows = []
    for idx, item in enumerate(line_items, start=1):
        if not isinstance(item, dict):
            continue
        rows.append({
            "Source Doc": d["filename"],
            "Doc Number": doc_number,
            "Line#": idx,
            "Part Number": item.get("part_number") or "",
            "Description": item.get("description") or "",
            "Qty": normalize_amount(item.get("quantity")),
            "Unit Price": normalize_amount(item.get("unit_price")),
            "Tax": normalize_amount(item.get("tax")),
            "Line Amount": normalize_amount(item.get("total")),
        })
    return rows


def _build_line_items_sheet(wb: Workbook, documents: list[dict]) -> None:
    ws = wb.create_sheet("Line Items")
    rows = [[row[h] for h in LINE_ITEM_COLUMNS] for d in _ok_docs(documents) for row in line_item_rows(d)]

    _write_table(ws, 1, LINE_ITEM_COLUMNS, rows)
    ws.freeze_panes = "A2"
    _autosize_columns(ws, len(LINE_ITEM_COLUMNS))


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def build_excel_report(documents: list[dict], custom_field_names: list[str] | None = None) -> BytesIO:
    """
    Build the 3-sheet DocuSense AI Excel report: Summary Dashboard, Header
    Details, Line Items.

    documents: list of {filename, doc_type, success, data, error}
    custom_field_names: optional user-defined field names to add as extra columns in "Header Details"
    """
    wb = Workbook()
    wb.remove(wb.active)  # drop the default blank sheet; we add our own in order

    generated_at = datetime.now()

    _build_summary_dashboard(wb, documents, generated_at)
    _build_header_details_sheet(wb, documents, custom_field_names=custom_field_names)
    _build_line_items_sheet(wb, documents)

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer
