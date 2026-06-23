"""
DocuSense AI — uipath_bridge.py

Entry point for calling this project's extraction + Excel/DataTable-row
logic directly from external callers, with no Streamlit dependency.
Used by uipath_cli.py (the recommended path for UiPath — see that file),
and equally callable from UiPath's "Invoke Python Method" activity
(UiPath.Python.Activities) or any other plain Python caller.

Every function takes and returns plain strings (file paths, JSON) rather
than Python objects, since that's what marshals cleanly regardless of
which calling mechanism is used.
"""

from __future__ import annotations

import json
import os

from excel_builder import build_excel_report, header_detail_row, line_item_rows
from extractor import DEFAULT_MODEL_NAME, configure_api, extract_document


def _extract_batch(api_key: str, file_paths: list, model_name: str, doc_type: str) -> list:
    resolved_model_name = model_name.strip() or DEFAULT_MODEL_NAME
    model = configure_api(api_key, resolved_model_name)
    results = []
    for path in file_paths:
        filename = os.path.basename(path)
        try:
            with open(path, "rb") as f:
                file_bytes = f.read()
            result = extract_document(model, filename, file_bytes, doc_type, primary_model_name=resolved_model_name)
        except Exception as e:
            result = {"filename": filename, "doc_type": doc_type, "success": False, "data": None, "error": str(e)}
        results.append(result)
    return results


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
        results = _extract_batch(api_key, file_paths, model_name, doc_type)
    except Exception as e:
        return json.dumps({"success": False, "error": f"Setup failed: {e}"})

    try:
        buf = build_excel_report(results)
        with open(output_excel_path, "wb") as f:
            f.write(buf.getvalue())
    except Exception as e:
        return json.dumps({"success": False, "error": f"Excel report build failed: {e}"})

    summary = [
        {"filename": r["filename"], "success": r["success"], "error": r["error"], "model_used": r.get("model_used")}
        for r in results
    ]
    return json.dumps({"success": True, "excel_path": output_excel_path, "results": summary})


def process_documents_to_rows(
    api_key: str,
    file_paths_json: str,
    model_name: str = "",
    doc_type: str = "Invoice",
) -> str:
    """
    Extracts every file listed in file_paths_json and returns flat JSON rows — one row per
    document for headers, one row per line item — ready to convert directly into UiPath
    DataTables (e.g. via "Deserialize JSON Array" / building a DataTable from the array),
    which you can then write to Excel yourself with UiPath's own Excel activities:

        {
          "success": true,
          "header_rows": [{"File Name": ..., "Document Type": ..., ...}, ...],
          "line_item_rows": [{"Source Doc": ..., "Doc Number": ..., ...}, ...],
          "errors": [{"filename": ..., "error": ...}, ...]
        }
        {"success": false, "error": "..."}

    Use this instead of process_documents_to_excel() when you want the data back as UiPath
    DataTables rather than an already-built Excel file.
    """
    try:
        file_paths = json.loads(file_paths_json)
        if not isinstance(file_paths, list) or not file_paths:
            return json.dumps({"success": False, "error": "file_paths_json must be a non-empty JSON array of file paths."})
        results = _extract_batch(api_key, file_paths, model_name, doc_type)
    except Exception as e:
        return json.dumps({"success": False, "error": f"Setup failed: {e}"})

    header_rows = [header_detail_row(r) for r in results]
    item_rows = [row for r in results for row in line_item_rows(r)]
    errors = [{"filename": r["filename"], "error": r["error"]} for r in results if not r["success"]]

    return json.dumps({"success": True, "header_rows": header_rows, "line_item_rows": item_rows, "errors": errors})


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
        resolved_model_name = model_name.strip() or DEFAULT_MODEL_NAME
        model = configure_api(api_key, resolved_model_name)
        with open(file_path, "rb") as f:
            file_bytes = f.read()
        result = extract_document(model, filename, file_bytes, doc_type, primary_model_name=resolved_model_name)
    except Exception as e:
        return json.dumps({"success": False, "filename": filename, "error": str(e)})

    if not result["success"]:
        return json.dumps({"success": False, "filename": filename, "error": result["error"]})
    return json.dumps({"success": True, "filename": filename, "data": result["data"], "model_used": result.get("model_used")})
