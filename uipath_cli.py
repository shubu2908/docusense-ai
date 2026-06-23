"""
DocuSense AI — uipath_cli.py

Command-line entry point for calling uipath_bridge.py from UiPath via a
plain "Invoke Power Shell" / "Start Process" activity — no special Python
activity package required. Always prints exactly one line of JSON to
stdout. Exit code is 0 if the script itself ran (check the JSON's
"success" field for the actual extraction outcome); 1 only for invalid
arguments or an unexpected crash before any JSON could be produced.

Usage:
  python uipath_cli.py to-excel --api-key KEY --files "a.pdf|b.tif" --output out.xlsx [--model NAME] [--doc-type Invoice] [--required-fields "po_number,total"] [--result-file result.json]
  python uipath_cli.py to-rows  --api-key KEY --files "a.pdf|b.tif" [--model NAME] [--doc-type Invoice] [--required-fields "po_number,total"] [--result-file result.json]
  python uipath_cli.py single   --api-key KEY --file  a.pdf         [--model NAME] [--doc-type Invoice] [--required-fields "po_number,total"] [--result-file result.json]

--files takes a "|"-delimited list of absolute file paths (build it in UiPath with
String.Join("|", arrayOfPaths)) since that's simpler to construct from a workflow than JSON.

--required-fields takes a comma-separated list of field names that must be present - missing
ones (plus any field Gemini itself reports "Low" confidence in, and a Subtotal+Tax+Shipping
vs Total mismatch) come back as "Validation Status"/"Validation Notes" (to-excel/to-rows) or
"validation_issues" (single).

--result-file, if given, also writes the JSON result to that path. Prefer this over
capturing stdout from "Start Process"/"Invoke Power Shell" in UiPath — the exact
stdout-capture property differs across UiPath package versions, but "Start Process" +
"Read Text File" on a known path works identically everywhere.
"""

from __future__ import annotations

import argparse
import json
import sys

from uipath_bridge import (
    extract_single_document,
    process_documents_to_excel,
    process_documents_to_rows,
)


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="uipath_cli.py")
    sub = parser.add_subparsers(dest="command", required=True)

    p_excel = sub.add_parser("to-excel")
    p_excel.add_argument("--api-key", required=True)
    p_excel.add_argument("--files", required=True, help='"|"-separated list of absolute file paths')
    p_excel.add_argument("--output", required=True, help="Absolute path to write the .xlsx report to")
    p_excel.add_argument("--model", default="")
    p_excel.add_argument("--doc-type", default="Invoice")
    p_excel.add_argument("--required-fields", default="")
    p_excel.add_argument("--result-file", default="")

    p_rows = sub.add_parser("to-rows")
    p_rows.add_argument("--api-key", required=True)
    p_rows.add_argument("--files", required=True, help='"|"-separated list of absolute file paths')
    p_rows.add_argument("--model", default="")
    p_rows.add_argument("--doc-type", default="Invoice")
    p_rows.add_argument("--required-fields", default="")
    p_rows.add_argument("--result-file", default="")

    p_single = sub.add_parser("single")
    p_single.add_argument("--api-key", required=True)
    p_single.add_argument("--file", required=True)
    p_single.add_argument("--model", default="")
    p_single.add_argument("--doc-type", default="Invoice")
    p_single.add_argument("--required-fields", default="")
    p_single.add_argument("--result-file", default="")

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = _build_parser()
    try:
        args = parser.parse_args(argv)
    except SystemExit:
        # argparse already wrote a usage error to stderr; stdout must stay JSON-only.
        print(json.dumps({"success": False, "error": "Invalid command-line arguments."}))
        return 1

    try:
        if args.command == "to-excel":
            files = [f for f in args.files.split("|") if f]
            result = process_documents_to_excel(
                api_key=args.api_key,
                file_paths_json=json.dumps(files),
                output_excel_path=args.output,
                model_name=args.model,
                doc_type=args.doc_type,
                required_fields_csv=args.required_fields,
            )
        elif args.command == "to-rows":
            files = [f for f in args.files.split("|") if f]
            result = process_documents_to_rows(
                api_key=args.api_key,
                file_paths_json=json.dumps(files),
                model_name=args.model,
                doc_type=args.doc_type,
                required_fields_csv=args.required_fields,
            )
        else:  # single
            result = extract_single_document(
                api_key=args.api_key,
                file_path=args.file,
                model_name=args.model,
                doc_type=args.doc_type,
                required_fields_csv=args.required_fields,
            )
    except Exception as e:
        result = json.dumps({"success": False, "error": f"Unexpected error: {e}"})
        print(result)
        if args.result_file:
            _write_result_file(args.result_file, result)
        return 1

    print(result)
    if args.result_file:
        _write_result_file(args.result_file, result)
    return 0


def _write_result_file(path: str, content: str) -> None:
    try:
        with open(path, "w", encoding="utf-8") as f:
            f.write(content)
    except Exception as e:
        # The JSON result is already on stdout; a failed file write shouldn't mask that.
        print(json.dumps({"success": False, "error": f"Could not write --result-file '{path}': {e}"}), file=sys.stderr)


if __name__ == "__main__":
    sys.exit(main())
