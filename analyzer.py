"""
DocuSense AI — analyzer.py

Cross-document intelligence.

Numeric/string matching (PO<->Invoice matching, >5% mismatch flags,
duplicate document numbers, contract expiry, vendor spend totals) is
computed deterministically in Python — these are exact arithmetic/string
tasks that a rule-based pass handles perfectly and for free, with none of
an LLM's rounding or hallucination risk.

On top of that, a single Gemini call is made with all extracted document
data to surface qualitative issues a fixed rule set can't anticipate
(odd payment terms, risky clauses, inconsistent statuses, etc.). If that
call fails for any reason, the deterministic results are still returned
in full — cross-document intelligence never goes blank just because the
AI enrichment step had a bad day.
"""

from __future__ import annotations

import json
import re
from datetime import datetime
from typing import Any, Optional

import pandas as pd

from extractor import call_gemini_with_json_retry

EXPIRY_WARNING_DAYS = 30
MISMATCH_THRESHOLD_PCT = 5.0


# ---------------------------------------------------------------------------
# Small shared helpers
# ---------------------------------------------------------------------------

def normalize_amount(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    text = re.sub(r"[^0-9.\-]", "", text)
    if text in ("", "-", "."):
        return None
    try:
        return float(text)
    except ValueError:
        return None


def normalize_key(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[^A-Za-z0-9]", "", str(value)).upper()


def normalize_name(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip().lower()


def parse_date(value: Any) -> Optional[pd.Timestamp]:
    if not value or isinstance(value, (list, dict)):
        return None
    try:
        ts = pd.to_datetime(value, errors="coerce")
    except (TypeError, ValueError):
        return None
    if not isinstance(ts, pd.Timestamp) or pd.isna(ts):
        return None
    return ts


def get_doc_amount(doc_type: str, data: dict) -> Optional[float]:
    if doc_type == "Invoice":
        return normalize_amount(data.get("total") if data.get("total") is not None else data.get("total_amount"))
    if doc_type == "Purchase Order":
        return normalize_amount(data.get("total_value") if data.get("total_value") is not None else data.get("total_amount"))
    if doc_type == "Contract":
        return normalize_amount(data.get("value") if data.get("value") is not None else data.get("total_amount"))
    return normalize_amount(data.get("total_amount"))


def get_doc_number(doc_type: str, data: dict) -> Optional[str]:
    key_by_type = {
        "Invoice": "invoice_number",
        "Purchase Order": "po_number",
        "Contract": "contract_id",
    }
    specific_key = key_by_type.get(doc_type)
    if specific_key and data.get(specific_key):
        return str(data[specific_key])
    return str(data["document_number"]) if data.get("document_number") else None


def get_vendor(doc_type: str, data: dict) -> Optional[str]:
    if doc_type in ("Invoice", "Purchase Order") and data.get("vendor"):
        return str(data["vendor"])
    return data.get("party_name_1") or None


def _ok_docs(documents: list[dict]) -> list[dict]:
    return [d for d in documents if d.get("success") and isinstance(d.get("data"), dict)]


# ---------------------------------------------------------------------------
# Deterministic: PO <-> Invoice matching + mismatch flags + missing docs
# ---------------------------------------------------------------------------

def find_po_invoice_matches(documents: list[dict]) -> tuple[list[dict], list[dict]]:
    """Returns (matches, issues)."""
    docs = _ok_docs(documents)
    pos = [d for d in docs if d["doc_type"] == "Purchase Order"]
    invoices = [d for d in docs if d["doc_type"] == "Invoice"]

    po_by_key = {}
    for po in pos:
        key = normalize_key(po["data"].get("po_number"))
        if key:
            po_by_key.setdefault(key, po)

    matches: list[dict] = []
    issues: list[dict] = []
    matched_po_keys: set[str] = set()

    for inv in invoices:
        inv_data = inv["data"]
        po_ref_key = normalize_key(inv_data.get("po_reference"))
        inv_number = get_doc_number("Invoice", inv_data) or inv["filename"]
        inv_amount = get_doc_amount("Invoice", inv_data)

        if not po_ref_key:
            continue  # invoice doesn't reference a PO at all — nothing to match

        po = po_by_key.get(po_ref_key)
        if po is None:
            issues.append({
                "severity": "Medium",
                "issue_type": "PO Reference Not Found",
                "documents_affected": inv["filename"],
                "details": f"Invoice {inv_number} references PO '{inv_data.get('po_reference')}' which was not found among uploaded documents.",
                "recommended_action": "Verify the PO number on the invoice, or upload the missing purchase order.",
            })
            continue

        matched_po_keys.add(po_ref_key)
        po_data = po["data"]
        po_number = get_doc_number("Purchase Order", po_data) or po["filename"]
        po_amount = get_doc_amount("Purchase Order", po_data)

        if po_amount is not None and inv_amount is not None and po_amount != 0:
            diff = abs(po_amount - inv_amount)
            diff_pct = (diff / abs(po_amount)) * 100
        elif po_amount is not None and inv_amount is not None:
            diff = abs(po_amount - inv_amount)
            diff_pct = 0.0 if diff == 0 else 100.0
        else:
            diff = None
            diff_pct = None

        if diff_pct is None:
            match_status, flag, severity = "Unverifiable", "Amount missing", None
        elif diff_pct > MISMATCH_THRESHOLD_PCT:
            match_status = "Mismatch"
            flag = f"{diff_pct:.1f}% difference"
            severity = "High" if diff_pct > 25 else "Medium"
        else:
            match_status, flag, severity = "Match", "", None

        matches.append({
            "po_number": po_number,
            "po_amount": po_amount,
            "invoice_number": inv_number,
            "invoice_amount": inv_amount,
            "match_status": match_status,
            "difference": round(diff, 2) if diff is not None else None,
            "flag": flag,
        })

        if severity:
            issues.append({
                "severity": severity,
                "issue_type": "Amount Mismatch",
                "documents_affected": f"{po['filename']} / {inv['filename']}",
                "details": f"PO {po_number} amount ({po_amount}) differs from Invoice {inv_number} amount ({inv_amount}) by {diff_pct:.1f}%.",
                "recommended_action": "Review the invoice against the purchase order and confirm the correct amount before payment.",
            })

    for key, po in po_by_key.items():
        if key not in matched_po_keys:
            po_number = get_doc_number("Purchase Order", po["data"]) or po["filename"]
            issues.append({
                "severity": "Medium",
                "issue_type": "Missing Invoice for PO",
                "documents_affected": po["filename"],
                "details": f"Purchase Order {po_number} has no matching invoice among uploaded documents.",
                "recommended_action": "Confirm whether goods/services were delivered and request the corresponding invoice.",
            })

    return matches, issues


# ---------------------------------------------------------------------------
# Deterministic: duplicate document numbers
# ---------------------------------------------------------------------------

def detect_duplicates(documents: list[dict]) -> list[dict]:
    docs = _ok_docs(documents)
    groups: dict[tuple[str, str], list[dict]] = {}

    for d in docs:
        number = get_doc_number(d["doc_type"], d["data"])
        key = normalize_key(number)
        if not key:
            continue
        groups.setdefault((d["doc_type"], key), []).append(d)

    issues = []
    for (doc_type, _key), group in groups.items():
        if len(group) < 2:
            continue
        filenames = ", ".join(g["filename"] for g in group)
        doc_number = get_doc_number(doc_type, group[0]["data"])
        issues.append({
            "severity": "High",
            "issue_type": "Duplicate Document Number",
            "documents_affected": filenames,
            "details": f"{len(group)} {doc_type} documents share the same document number ('{doc_number}').",
            "recommended_action": "Confirm these aren't duplicate submissions of the same document before processing payment or approval.",
        })
    return issues


# ---------------------------------------------------------------------------
# Deterministic: expired / expiring contracts
# ---------------------------------------------------------------------------

def detect_expiring_contracts(documents: list[dict], today: Optional[datetime] = None) -> list[dict]:
    today_ts = pd.Timestamp(today or datetime.now()).normalize()
    docs = [d for d in _ok_docs(documents) if d["doc_type"] == "Contract"]

    issues = []
    for d in docs:
        expiry = parse_date(d["data"].get("expiry_date"))
        if expiry is None:
            continue
        days_left = (expiry.normalize() - today_ts).days
        contract_id = get_doc_number("Contract", d["data"]) or d["filename"]

        if days_left < 0:
            issues.append({
                "severity": "High",
                "issue_type": "Contract Expired",
                "documents_affected": d["filename"],
                "details": f"Contract {contract_id} expired on {expiry.date()} ({abs(days_left)} days ago).",
                "recommended_action": "Renew, renegotiate, or formally close out this contract.",
            })
        elif days_left <= EXPIRY_WARNING_DAYS:
            issues.append({
                "severity": "Medium",
                "issue_type": "Contract Expiring Soon",
                "documents_affected": d["filename"],
                "details": f"Contract {contract_id} expires on {expiry.date()} (in {days_left} days).",
                "recommended_action": "Begin renewal discussions or plan for contract end before the expiry date.",
            })
    return issues


# ---------------------------------------------------------------------------
# Deterministic: vendor spend summary
# ---------------------------------------------------------------------------

def vendor_spend_summary(documents: list[dict]) -> list[dict]:
    docs = [d for d in _ok_docs(documents) if d["doc_type"] in ("Invoice", "Purchase Order")]
    totals: dict[tuple[str, str], dict] = {}

    for d in docs:
        vendor = get_vendor(d["doc_type"], d["data"])
        if not vendor:
            continue
        currency = (d["data"].get("currency") or "—").strip() or "—"
        amount = get_doc_amount(d["doc_type"], d["data"]) or 0.0
        key = (normalize_name(vendor), currency)

        row = totals.setdefault(key, {
            "vendor": vendor,
            "currency": currency,
            "invoice_total": 0.0,
            "po_total": 0.0,
            "document_count": 0,
        })
        row["document_count"] += 1
        if d["doc_type"] == "Invoice":
            row["invoice_total"] += amount
        else:
            row["po_total"] += amount

    rows = list(totals.values())
    rows.sort(key=lambda r: r["invoice_total"] + r["po_total"], reverse=True)
    return rows


# ---------------------------------------------------------------------------
# Gemini qualitative enrichment pass
# ---------------------------------------------------------------------------

def _build_compact_summary(documents: list[dict]) -> list[dict]:
    """Trim each extracted doc down to the fields useful for cross-doc reasoning, to keep the prompt small."""
    compact = []
    for d in _ok_docs(documents):
        data = d["data"]
        compact.append({
            "filename": d["filename"],
            "doc_type": d["doc_type"],
            "document_number": get_doc_number(d["doc_type"], data),
            "party_1": data.get("party_name_1"),
            "party_2": data.get("party_name_2"),
            "vendor": get_vendor(d["doc_type"], data),
            "amount": get_doc_amount(d["doc_type"], data),
            "currency": data.get("currency"),
            "status": data.get("status"),
            "document_date": data.get("document_date"),
            "po_reference": data.get("po_reference"),
            "payment_terms": data.get("payment_terms"),
            "due_date": data.get("due_date"),
            "expiry_date": data.get("expiry_date"),
            "renewal_terms": data.get("renewal_terms"),
            "penalties": data.get("penalties"),
            "approval_status": data.get("approval_status"),
        })
    return compact


def build_cross_doc_prompt(documents: list[dict]) -> str:
    compact = _build_compact_summary(documents)
    today = datetime.now().strftime("%Y-%m-%d")

    return f"""You are a financial document auditor reviewing a batch of already-extracted documents.

Today's date is {today}.

Here is the extracted data for every document in this batch (JSON array):
{json.dumps(compact, indent=2, default=str)}

Deterministic checks (PO/invoice matching, >5% amount mismatches, duplicate numbers, contract expiry) have
already been run separately. Your job is to find ADDITIONAL issues a fixed rule set would miss — for example:
inconsistent or suspicious statuses, unusual payment terms, vague or one-sided contract clauses, missing
critical fields that should be present given the document type, or anything else that looks operationally risky.

Return ONLY a single valid JSON object with this exact shape, and nothing else:
{{
  "issues": [
    {{
      "severity": "High" | "Medium" | "Low",
      "issue_type": "short label",
      "documents_affected": "filename(s) involved, comma-separated",
      "details": "1-2 sentence explanation",
      "recommended_action": "1 short sentence"
    }}
  ]
}}

If you find nothing noteworthy, return {{"issues": []}}. Do not repeat amount-mismatch, duplicate-number, or
contract-expiry findings — those are already handled elsewhere. No markdown, no commentary, JSON only.
"""


def gemini_qualitative_issues(model, documents: list[dict]) -> tuple[list[dict], Optional[str]]:
    if not _ok_docs(documents):
        return [], None

    prompt = build_cross_doc_prompt(documents)
    data, error = call_gemini_with_json_retry(model, prompt)

    if data is None:
        return [], error

    issues = data.get("issues") if isinstance(data, dict) else None
    if not isinstance(issues, list):
        return [], "Gemini response did not contain an 'issues' array."

    cleaned = []
    for item in issues:
        if not isinstance(item, dict):
            continue
        severity = str(item.get("severity") or "Low").title()
        if severity not in ("High", "Medium", "Low"):
            severity = "Low"
        cleaned.append({
            "severity": severity,
            "issue_type": item.get("issue_type") or "AI-Detected Issue",
            "documents_affected": item.get("documents_affected") or "",
            "details": item.get("details") or "",
            "recommended_action": item.get("recommended_action") or "",
        })
    return cleaned, None


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def run_cross_document_analysis(model, documents: list[dict]) -> dict:
    """
    Runs the full cross-document intelligence pass.

    Returns:
        {
            "matches": [...],
            "issues": [...],
            "vendor_spend": [...],
            "ai_enrichment_success": bool,
            "ai_error": str | None,
        }
    """
    matches, match_issues = find_po_invoice_matches(documents)
    duplicate_issues = detect_duplicates(documents)
    expiry_issues = detect_expiring_contracts(documents)
    vendor_spend = vendor_spend_summary(documents)

    issues = match_issues + duplicate_issues + expiry_issues

    ai_issues: list[dict] = []
    ai_error: Optional[str] = None
    try:
        ai_issues, ai_error = gemini_qualitative_issues(model, documents)
    except Exception as e:
        ai_error = f"AI enrichment failed: {e}"

    issues.extend(ai_issues)

    severity_order = {"High": 0, "Medium": 1, "Low": 2}
    issues.sort(key=lambda i: severity_order.get(i.get("severity"), 3))

    return {
        "matches": matches,
        "issues": issues,
        "vendor_spend": vendor_spend,
        "ai_enrichment_success": ai_error is None,
        "ai_error": ai_error,
    }
