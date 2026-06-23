"""
DocuSense AI — extractor.py

Handles all direct communication with Gemini:
- Building extraction prompts from one fixed field schema (FIELDS)
- Sending PDF/image bytes to Gemini for structured extraction
- Defensive JSON parsing with a single automatic retry on failure

This module also exposes a couple of small JSON-handling utilities
(`clean_json_text`, `safe_json_loads`) and a generic
`call_gemini_with_json_retry` helper that analyzer.py reuses for its
(currently unwired) cross-document intelligence pass, so there is
exactly one place that knows how to coax Gemini into returning clean JSON.
"""

from __future__ import annotations

import io
import json
import re
import time
from typing import Any, Optional

import google.generativeai as genai
from PIL import Image

# Google periodically retires/renames Gemini model versions, so these are sensible
# defaults, not hardcoded requirements — app.py exposes the choice as a sidebar dropdown
# (with a manual-entry escape hatch), and each one has its own separate free-tier quota,
# which is what makes the cross-model fallback in extract_document() useful.
KNOWN_MODELS = [
    "gemini-3.5-flash",
    "gemini-2.5-flash",
    "gemini-2.5-flash-lite",
    "gemini-2.0-flash",
    "gemini-2.0-flash-lite",
    "gemini-flash-latest",
]
DEFAULT_MODEL_NAME = KNOWN_MODELS[0]

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
# Field schema
#
# One fixed schema is used for every document regardless of the doc_type
# label assigned in Step 2 — that label is still sent to Gemini as context,
# but it no longer changes which fields get requested.
# ---------------------------------------------------------------------------

FIELDS = {
    "document_type": "Best-guess type of this document: Invoice, Contract, Purchase Order, Report, or Other",
    "invoice_number": "The invoice number/ID",
    "invoice_date": "The date the invoice was issued, formatted YYYY-MM-DD if determinable",
    "po_number": "Any purchase order number referenced on the document, or null",
    "vendor_name": "The vendor/seller issuing the document",
    "bill_to": "The party being billed",
    "line_items": (
        "Array of items, each {\"part_number\": str, \"description\": str, \"quantity\": number, "
        "\"unit_price\": number, \"tax\": number, \"total\": number}"
    ),
    "subtotal": "Subtotal amount before tax and shipping, as a plain number",
    "tax": "Total tax amount, as a plain number",
    "shipping_charges": "Shipping/freight/handling charges, as a plain number, or null",
    "total": "Final total amount, as a plain number",
    "payment_terms": "Payment terms text (e.g. 'Net 30')",
    "due_date": "Payment due date, formatted YYYY-MM-DD if determinable",
    "bank_details": "Any bank/payment account details listed (account number, IFSC/SWIFT, bank name), as a single string or null",
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

DEFAULT_LIST_FIELDS = {"line_items"}


def get_all_field_names() -> list[str]:
    """All built-in field names, for building a 'remove field' picker."""
    return sorted(FIELDS)


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
    all_fields = dict(FIELDS)

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
        candidate = response.candidates[0]
        finish_reason = getattr(candidate, "finish_reason", None)
        safety_ratings = getattr(candidate, "safety_ratings", None)
        raise RuntimeError(
            f"Could not read text from Gemini response "
            f"(finish_reason={finish_reason}, safety_ratings={safety_ratings}): {e}"
        ) from e


MAX_RATE_LIMIT_RETRIES = 2
DEFAULT_RATE_LIMIT_DELAY = 20.0
_RETRY_DELAY_RE = re.compile(r"retry_delay\s*\{?\s*seconds:\s*(\d+)", re.IGNORECASE)


def _is_rate_limit_error(exc: Exception) -> bool:
    text = str(exc)
    return "429" in text or "ResourceExhausted" in type(exc).__name__ or "quota" in text.lower()


def _is_daily_quota_error(exc: Exception) -> bool:
    """PerDay quota errors won't resolve by waiting a few seconds, unlike PerMinute ones - don't retry these."""
    return "perday" in str(exc).lower().replace(" ", "")


_AUTH_ERROR_MARKERS = ("api key not valid", "api_key_invalid", "401", "403", "permission_denied", "unauthenticated")


def _is_auth_error(exc: Exception) -> bool:
    """A bad/invalid API key fails identically for every model - no point cascading through fallbacks for this."""
    text = str(exc).lower()
    return any(marker in text for marker in _AUTH_ERROR_MARKERS)


def _extract_retry_delay(exc: Exception) -> float:
    match = _RETRY_DELAY_RE.search(str(exc))
    return float(match.group(1)) + 1 if match else DEFAULT_RATE_LIMIT_DELAY


def _send_with_rate_limit_retry(chat, message, generation_config) -> tuple[Optional[Any], Optional[Exception]]:
    """Sends a chat message, automatically waiting and retrying on a per-minute 429 rate-limit error.
    Per-day quota errors are not retried, since waiting a few seconds can never resolve those."""
    last_error: Optional[Exception] = None
    for attempt in range(MAX_RATE_LIMIT_RETRIES + 1):
        try:
            return chat.send_message(message, generation_config=generation_config), None
        except Exception as e:
            last_error = e
            if _is_rate_limit_error(e) and not _is_daily_quota_error(e) and attempt < MAX_RATE_LIMIT_RETRIES:
                time.sleep(_extract_retry_delay(e))
                continue
            return None, e
    return None, last_error


_DAILY_QUOTA_MARKER = "DAILY quota for this model is exhausted"


def _friendly_error_message(prefix: str, error: Exception) -> str:
    if _is_daily_quota_error(error):
        return f"{prefix}: {error}\n\n{_DAILY_QUOTA_MARKER} (resets ~24h after first use today)."
    return f"{prefix}: {error}"


def call_gemini_with_json_retry(model: "genai.GenerativeModel", first_message: list | str) -> tuple[Optional[Any], Optional[str]]:
    """
    Send `first_message` to Gemini expecting JSON back. If parsing fails, send one
    follow-up message (in the same chat session) asking for cleaner JSON. Rate-limit
    (429) errors are retried automatically, waiting whatever delay Gemini suggests.

    Returns (data, error). data is None if both attempts failed.
    """
    generation_config = genai.types.GenerationConfig(
        temperature=0.1,
        response_mime_type="application/json",
    )

    chat = model.start_chat()
    response, error = _send_with_rate_limit_retry(chat, first_message, generation_config)
    if error is not None:
        return None, _friendly_error_message("Gemini request failed", error)
    try:
        text = _extract_response_text(response)
    except Exception as e:
        return None, f"Gemini request failed: {e}"

    data, err = safe_json_loads(text)
    if data is not None:
        return data, None

    response2, error2 = _send_with_rate_limit_retry(chat, JSON_RETRY_INSTRUCTION, generation_config)
    if error2 is not None:
        return None, _friendly_error_message("Gemini retry request failed", error2)
    try:
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
    primary_model_name: str = "",
    fallback_models: list[str] | None = None,
) -> dict:
    """
    Extract structured data from a single document using Gemini.

    custom_fields: optional list of {"name", "description", "scope", "is_list"} to extract in addition
                   to the built-in schema (scope is "All" or one of DOC_TYPES).
    excluded_fields: optional set of built-in field names to skip extracting entirely.
    primary_model_name: the name `model` was constructed with, used only to label results and to
                   avoid re-trying the same model as its own fallback.
    fallback_models: models to try, in order, if the primary model's free-tier DAILY quota is
                   exhausted (each model has a separate quota). Defaults to KNOWN_MODELS minus
                   primary_model_name. Per-minute rate limits are already retried within a single
                   model by call_gemini_with_json_retry and never trigger a model switch.

    Returns a dict: {filename, doc_type, success, data, error, model_used}
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

    if fallback_models is None:
        fallback_models = [m for m in KNOWN_MODELS if m != primary_model_name]

    candidates = [(primary_model_name or "selected model", model)]
    for name in fallback_models:
        try:
            candidates.append((name, genai.GenerativeModel(name)))
        except Exception:
            continue  # name not constructible (e.g. retired) - just skip it, not fatal

    tried_names = []
    last_error = None
    for name, candidate_model in candidates:
        tried_names.append(name)
        data, error = call_gemini_with_json_retry(candidate_model, [prompt, *file_parts])

        if data is None:
            last_error = error
            if _is_auth_error(error):
                break  # a bad API key fails identically on every model - no point trying the rest
            continue  # quota, transient, or content-generation issues might not affect a different model

        if not isinstance(data, dict):
            last_error = f"Expected a JSON object but got {type(data).__name__}."
            continue

        result["success"] = True
        result["data"] = data
        result["model_used"] = name
        return result

    if len(tried_names) > 1:
        result["error"] = (
            f"All {len(tried_names)} models tried ({', '.join(tried_names)}) failed. Last error: {last_error}"
        )
    else:
        result["error"] = last_error or "Unknown extraction error."
    return result
