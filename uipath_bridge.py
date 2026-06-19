"""
DocuSense AI — uipath_bridge.py

Entry point for calling this project's extraction + Excel-building logic
directly from external callers (built for UiPath's "Invoke Python Method"
activity via the UiPath.Python.Activities package, but usable from any
plain Python caller — Streamlit is never imported here).

Every function takes and returns plain strings (file paths, JSON) rather
than Python objects, since that's what marshals cleanly across UiPath's
Python bridge regardless of activity-pack version.
"""

from __future__ import annotations

import json
import os

from excel_builder import build_excel_report
from extractor import DEFAULT_MODEL_NAME, configure_api, extract_document


def process_documents_to_excel(
    api_key: str,
    file_paths_json: str,
    output_excel_path: str,
    model_name: str = "",
    doc_type: str = "Invoice",
) -> str:
    """
    Extracts every file listed in file_paths_json (a JSON array of absolute file paths) via
    Gemini, writes a 3-sheet Excel report (Summary Dashboard, Header Details, Line Items) to
    output_excel_path, and returns a JSON string describing the outcome:

        {"success": true, "excel_path": "...", "results": [{"filename":, "success":, "error":}, ...]}
        {"success": false, "error": "..."}

    One file failing extraction does not stop the batch — it shows up with "success": false
    in "results" and the Excel report still builds from whatever succeeded.
    """
    try:
        file_paths = json.loads(file_paths_json)
        if not isinstance(file_paths, list) or not file_paths:
            return json.dumps({"success": False, "error": "file_paths_json must be a non-empty JSON array of file paths."})
        model = configure_api(api_key, model_name.strip() or DEFAULT_MODEL_NAME)
    except Exception as e:
        return json.dumps({"success": False, "error": f"Setup failed: {e}"})

    results = []
    for path in file_paths:
        filename = os.path.basename(path)
        try:
            with open(path, "rb") as f:
                file_bytes = f.read()
            result = extract_document(model, filename, file_bytes, doc_type)
        except Exception as e:
            result = {"filename": filename, "doc_type": doc_type, "success": False, "data": None, "error": str(e)}
        results.append(result)

    try:
        buf = build_excel_report(results)
        with open(output_excel_path, "wb") as f:
            f.write(buf.getvalue())
    except Exception as e:
        return json.dumps({"success": False, "error": f"Excel report build failed: {e}"})

    summary = [{"filename": r["filename"], "success": r["success"], "error": r["error"]} for r in results]
    return json.dumps({"success": True, "excel_path": output_excel_path, "results": summary})


def extract_single_document(api_key: str, file_path: str, model_name: str = "", doc_type: str = "Invoice") -> str:
    """
    Extracts one document and returns its structured data as a JSON string, with no Excel
    file produced — useful when a workflow wants the fields directly (e.g. to populate a
    form or compare against an ERP record):

        {"success": true, "filename": "...", "data": {...}}
        {"success": false, "filename": "...", "error": "..."}
    """
    filename = os.path.basename(file_path)
    try:
        model = configure_api(api_key, model_name.strip() or DEFAULT_MODEL_NAME)
        with open(file_path, "rb") as f:
            file_bytes = f.read()
        result = extract_document(model, filename, file_bytes, doc_type)
    except Exception as e:
        return json.dumps({"success": False, "filename": filename, "error": str(e)})

    if not result["success"]:
        return json.dumps({"success": False, "filename": filename, "error": result["error"]})
    return json.dumps({"success": True, "filename": filename, "data": result["data"]})
