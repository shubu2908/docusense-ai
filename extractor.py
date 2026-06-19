"""
DocuSense AI — extractor.py

Handles all direct communication with Google Gemini 1.5 Flash:
- Building document-type-aware extraction prompts
- Sending PDF/image bytes to Gemini for structured extraction
- Defensive JSON parsing with a single automatic retry on failure

This module also exposes a couple of small JSON-handling utilities
(`clean_json_text`, `safe_json_loads`) and a generic
`call_gemini_with_json_retry` helper that analyzer.py reuses for the
cross-document intelligence pass, so there is exactly one place that
knows how to coax Gemini into returning clean JSON.
"""

from __future__ import annotations

import io
import json
import re
from typing import Any, Optional

import google.generativeai as genai
from PIL import Image

# Google periodically retires old Gemini model versions, so this is a default,
# not a hardcoded requirement — app.py exposes it as an overridable sidebar setting.
DEFAULT_MODEL_NAME = "gemini-3.5-flash"

MIME_TYPES = {
    "pdf": "application/pdf",
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "tif": "image/tiff",
    "tiff": "image/tiff",
}

# Gemini has no native TIFF support, so TIFF pages are converted to PNG in-memory.
# A multi-page TIFF becomes one image part per page, all sent as the same document.
MAX_TIFF_PAGES = 30

# ---------------------------------------------------------------------------
# Field schemas
# ---------------------------------------------------------------------------

COMMON_FIELDS = {
    "document_type": "Best-guess type of this document: Invoice, Contract, Purchase Order, Report, or Other",
    "document_number": "The primary identifying number/ID printed on the document",
    "document_date": "The date the document was issued, formatted YYYY-MM-DD if determinable",
    "party_name_1": "The first named party (issuer, seller, vendor, disclosing party, etc.)",
    "party_name_2": "The second named party (recipient, buyer, client, receiving party, etc.)",
    "total_amount": "The headline total monetary amount on the document, as a plain number (no currency symbols, no commas)",
    "currency": "ISO currency code if determinable (e.g. USD, EUR, INR, GBP)",
    "status": "Any status shown on the document (e.g. Paid, Pending, Draft, Signed, Approved, Overdue)",
    "key_dates": "Array of other important dates found, each item shaped like {\"label\": str, \"date\": \"YYYY-MM-DD\"}",
    "summary": "A concise 1-3 sentence plain-English summary of what this document is and what it says",
}

INVOICE_FIELDS = {
    "invoice_number": "The invoice number/ID",
    "vendor": "The vendor/seller issuing the invoice",
    "bill_to": "The party being billed",
    "po_reference": "Any purchase order number referenced on the invoice, or null",
    "line_items": "Array of items, each {\"description\": str, \"quantity\": number, \"unit_price\": number, \"tax\": number, \"total\": number}",
    "subtotal": "Subtotal amount before tax, as a plain number",
    "tax": "Total tax amount, as a plain number",
    "total": "Final total amount, as a plain number",
    "payment_terms": "Payment terms text (e.g. 'Net 30')",
    "due_date": "Payment due date, formatted YYYY-MM-DD if determinable",
    "bank_details": "Any bank/payment account details listed (account number, IFSC/SWIFT, bank name), as a single string or null",
}

CONTRACT_FIELDS = {
    "contract_id": "The contract reference number/ID",
    "parties": "Array of strings naming all parties to the contract",
    "effective_date": "Contract start/effective date, formatted YYYY-MM-DD if determinable",
    "expiry_date": "Contract end/expiry date, formatted YYYY-MM-DD if determinable",
    "value": "Total contract value, as a plain number, or null",
    "jurisdiction": "Governing law / jurisdiction stated in the contract",
    "key_obligations": "Array of strings, each a short description of a key obligation",
    "termination_clause": "Short summary of the termination clause, or null",
    "renewal_terms": "Short summary of renewal/auto-renewal terms, or null",
    "penalties": "Short summary of any penalty/liquidated-damages clauses, or null",
}

PO_FIELDS = {
    "po_number": "The purchase order number",
    "vendor": "The vendor the PO was issued to",
    "delivery_date": "Expected delivery date, formatted YYYY-MM-DD if determinable",
    "line_items": "Array of items, each {\"description\": str, \"quantity\": number, \"unit_price\": number, \"tax\": number, \"total\": number}",
    "total_value": "Total PO value, as a plain number",
    "delivery_address": "Delivery address text, or null",
    "approval_status": "Approval status if shown (e.g. Approved, Pending Approval)",
}

TYPE_SPECIFIC_FIELDS = {
    "Invoice": INVOICE_FIELDS,
    "Contract": CONTRACT_FIELDS,
    "Purchase Order": PO_FIELDS,
}

DOC_TYPES = ["Invoice", "Contract", "Purchase Order", "Report", "Other"]


# ---------------------------------------------------------------------------
# Gemini setup
# ---------------------------------------------------------------------------

def configure_api(api_key: str, model_name: str = DEFAULT_MODEL_NAME) -> "genai.GenerativeModel":
    """Configure the Gemini client and return a ready-to-use model handle."""
    if not api_key or not api_key.strip():
        raise ValueError("A Gemini API key is required.")
    if not model_name or not model_name.strip():
        raise ValueError("A Gemini model name is required.")
    genai.configure(api_key=api_key.strip())
    return genai.GenerativeModel(model_name.strip())


def get_mime_type(filename: str) -> str:
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    if ext not in MIME_TYPES:
        raise ValueError(f"Unsupported file type: '.{ext}'. Allowed: PDF, PNG, JPG, JPEG, TIF, TIFF.")
    return MIME_TYPES[ext]


def _tiff_to_png_parts(file_bytes: bytes) -> list[dict]:
    """Converts each page of a (possibly multi-page) TIFF into a separate PNG image part."""
    parts = []
    with Image.open(io.BytesIO(file_bytes)) as img:
        n_frames = getattr(img, "n_frames", 1)
        for i in range(min(n_frames, MAX_TIFF_PAGES)):
            img.seek(i)
            buf = io.BytesIO()
            img.convert("RGB").save(buf, format="PNG")
            parts.append({"mime_type": "image/png", "data": buf.getvalue()})
    return parts


def build_file_parts(filename: str, file_bytes: bytes) -> list[dict]:
    """Builds the Gemini content part(s) for a file, converting TIFF to PNG since Gemini has no native TIFF support."""
    mime_type = get_mime_type(filename)
    if mime_type == "image/tiff":
        return _tiff_to_png_parts(file_bytes)
    return [{"mime_type": mime_type, "data": file_bytes}]


# ---------------------------------------------------------------------------
# Prompt building
# ---------------------------------------------------------------------------

DEFAULT_LIST_FIELDS = {"line_items", "parties", "key_obligations", "key_dates"}


def get_all_field_names() -> list[str]:
    """All field names known across the common + type-specific schemas, for building a 'remove field' picker."""
    names = set(COMMON_FIELDS)
    for fields in TYPE_SPECIFIC_FIELDS.values():
        names.update(fields)
    return sorted(names)


def sanitize_field_name(name: str) -> str:
    """Turns free-typed custom field names into safe snake_case-ish JSON keys."""
    cleaned = re.sub(r"\s+", "_", name.strip().lower())
    cleaned = re.sub(r"[^a-z0-9_]", "", cleaned)
    return cleaned.strip("_")


def _build_field_skeleton(fields: dict, list_field_names: set) -> dict:
    skeleton = {}
    for key in fields:
        skeleton[key] = [] if key in list_field_names else None
    return skeleton


def build_extraction_prompt(
    doc_type: str,
    custom_fields: list[dict] | None = None,
    excluded_fields: set | None = None,
) -> str:
    type_fields = TYPE_SPECIFIC_FIELDS.get(doc_type, {})
    all_fields = {**COMMON_FIELDS, **type_fields}

    if excluded_fields:
        all_fields = {k: v for k, v in all_fields.items() if k not in excluded_fields}

    list_field_names = set(DEFAULT_LIST_FIELDS)
    for cf in custom_fields or []:
        name = cf.get("name")
        if not name or cf.get("scope") not in ("All", doc_type):
            continue
        all_fields[name] = cf.get("description") or f"Custom field: {name}"
        if cf.get("is_list"):
            list_field_names.add(name)

    skeleton = _build_field_skeleton(all_fields, list_field_names)

    field_lines = "\n".join(f'- "{k}": {v}' for k, v in all_fields.items())

    prompt = f"""You are a precise document data-extraction engine used in a financial/legal document intelligence system.

The user has labeled this document as: {doc_type}

Carefully read the attached document and extract the following fields:

{field_lines}

STRICT OUTPUT RULES:
1. Return ONLY a single valid JSON object. No markdown, no code fences, no commentary, no explanations before or after.
2. Use exactly this set of keys (do not add or rename keys):
{json.dumps(skeleton, indent=2)}
3. If a field cannot be found in the document, use null for scalar fields and an empty array [] for list fields. Never invent data.
4. All monetary fields must be plain numbers (no currency symbols, no thousands separators, no quotes around numbers).
5. All dates must be formatted as YYYY-MM-DD whenever a full date can be determined.
6. Output must be parseable directly by a JSON parser with no post-processing.
"""
    return prompt


JSON_RETRY_INSTRUCTION = (
    "Your previous response could not be parsed as valid JSON. "
    "Resend your answer now as ONLY a single valid JSON object using the exact same fields and values you already determined. "
    "Do not include markdown formatting, code fences, backticks, comments, or any explanatory text — raw JSON only, starting with '{' and ending with '}'."
)


# ---------------------------------------------------------------------------
# JSON utilities
# ---------------------------------------------------------------------------

_CODE_FENCE_RE = re.compile(r"^```(?:json)?\s*|\s*```$", re.IGNORECASE | re.MULTILINE)


def clean_json_text(text: str) -> str:
    """Strip common Gemini formatting artifacts (code fences, stray prose) from a JSON-ish string."""
    cleaned = _CODE_FENCE_RE.sub("", text).strip()

    first_obj, first_arr = cleaned.find("{"), cleaned.find("[")
    candidates = [pos for pos in (first_obj, first_arr) if pos != -1]
    if not candidates:
        return cleaned
    start = min(candidates)

    last_obj, last_arr = cleaned.rfind("}"), cleaned.rfind("]")
    end = max(last_obj, last_arr)
    if end == -1 or end < start:
        return cleaned

    return cleaned[start : end + 1].strip()


def safe_json_loads(text: str) -> tuple[Optional[Any], Optional[str]]:
    """Try to parse text as JSON, falling back to a cleaned version. Returns (data, error)."""
    if not text:
        return None, "Empty response from model."
    try:
        return json.loads(text), None
    except json.JSONDecodeError:
        pass
    try:
        return json.loads(clean_json_text(text)), None
    except json.JSONDecodeError as e:
        return None, f"JSON parse error: {e}"


def _extract_response_text(response) -> str:
    if not getattr(response, "candidates", None):
        feedback = getattr(response, "prompt_feedback", None)
        raise RuntimeError(f"Gemini returned no candidates (possibly blocked). Feedback: {feedback}")
    try:
        return response.text
    except Exception as e:
        raise RuntimeError(f"Could not read text from Gemini response: {e}") from e


def call_gemini_with_json_retry(model: "genai.GenerativeModel", first_message: list | str) -> tuple[Optional[Any], Optional[str]]:
    """
    Send `first_message` to Gemini expecting JSON back. If parsing fails, send one
    follow-up message (in the same chat session) asking for cleaner JSON.

    Returns (data, error). data is None if both attempts failed.
    """
    generation_config = genai.types.GenerationConfig(
        temperature=0.1,
        response_mime_type="application/json",
    )

    chat = model.start_chat()
    try:
        response = chat.send_message(first_message, generation_config=generation_config)
        text = _extract_response_text(response)
    except Exception as e:
        return None, f"Gemini request failed: {e}"

    data, err = safe_json_loads(text)
    if data is not None:
        return data, None

    try:
        response2 = chat.send_message(JSON_RETRY_INSTRUCTION, generation_config=generation_config)
        text2 = _extract_response_text(response2)
    except Exception as e:
        return None, f"Gemini retry request failed: {e}"

    data2, err2 = safe_json_loads(text2)
    if data2 is not None:
        return data2, None

    return None, f"Could not parse JSON after retry: {err2 or err}"


# ---------------------------------------------------------------------------
# Public extraction entry point
# ---------------------------------------------------------------------------

def extract_document(
    model: "genai.GenerativeModel",
    filename: str,
    file_bytes: bytes,
    doc_type: str,
    custom_fields: list[dict] | None = None,
    excluded_fields: set | None = None,
) -> dict:
    """
    Extract structured data from a single document using Gemini.

    custom_fields: optional list of {"name", "description", "scope", "is_list"} to extract in addition
                   to the built-in schema (scope is "All" or one of DOC_TYPES).
    excluded_fields: optional set of built-in field names to skip extracting entirely.

    Returns a dict: {filename, doc_type, success, data, error}
    """
    result = {"filename": filename, "doc_type": doc_type, "success": False, "data": None, "error": None}

    if not file_bytes:
        result["error"] = "File is empty."
        return result

    try:
        file_parts = build_file_parts(filename, file_bytes)
    except ValueError as e:
        result["error"] = str(e)
        return result
    except Exception as e:
        result["error"] = f"Could not process file: {e}"
        return result

    prompt = build_extraction_prompt(doc_type, custom_fields=custom_fields, excluded_fields=excluded_fields)
    if len(file_parts) > 1:
        prompt += (
            f"\n\nNote: this document spans {len(file_parts)} pages, provided below as sequential "
            "page images. Treat them as one single document and combine information across all pages."
        )

    data, error = call_gemini_with_json_retry(model, [prompt, *file_parts])

    if data is None:
        result["error"] = error or "Unknown extraction error."
        return result

    if not isinstance(data, dict):
        result["error"] = f"Expected a JSON object but got {type(data).__name__}."
        return result

    result["success"] = True
    result["data"] = data
    return result
