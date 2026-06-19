"""
DocuSense AI — uipath_cli.py

Command-line entry point for calling uipath_bridge.py from UiPath via a
plain "Invoke Power Shell" / "Start Process" activity — no special Python
activity package required. Always prints exactly one line of JSON to
stdout. Exit code is 0 if the script itself ran (check the JSON's
"success" field for the actual extraction outcome); 1 only for invalid
arguments or an unexpected crash before any JSON could be produced.

Usage:
  python uipath_cli.py to-excel --api-key KEY --files "a.pdf|b.tif" --output out.xlsx [--model NAME] [--doc-type Invoice]
  python uipath_cli.py to-rows  --api-key KEY --files "a.pdf|b.tif" [--model NAME] [--doc-type Invoice]
  python uipath_cli.py single   --api-key KEY --file  a.pdf         [--model NAME] [--doc-type Invoice]

--files takes a "|"-delimited list of absolute file paths (build it in UiPath with
String.Join("|", arrayOfPaths)) since that's simpler to construct from a workflow than JSON.
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

    p_rows = sub.add_parser("to-rows")
    p_rows.add_argument("--api-key", required=True)
    p_rows.add_argument("--files", required=True, help='"|"-separated list of absolute file paths')
    p_rows.add_argument("--model", default="")
    p_rows.add_argument("--doc-type", default="Invoice")

    p_single = sub.add_parser("single")
    p_single.add_argument("--api-key", required=True)
    p_single.add_argument("--file", required=True)
    p_single.add_argument("--model", default="")
    p_single.add_argument("--doc-type", default="Invoice")

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
            )
        elif args.command == "to-rows":
            files = [f for f in args.files.split("|") if f]
            result = process_documents_to_rows(
                api_key=args.api_key,
                file_paths_json=json.dumps(files),
                model_name=args.model,
                doc_type=args.doc_type,
            )
        else:  # single
            result = extract_single_document(
                api_key=args.api_key,
                file_path=args.file,
                model_name=args.model,
                doc_type=args.doc_type,
            )
    except Exception as e:
        print(json.dumps({"success": False, "error": f"Unexpected error: {e}"}))
        return 1

    print(result)
    return 0


if __name__ == "__main__":
    sys.exit(main())
