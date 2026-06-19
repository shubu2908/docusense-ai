"""
DocuSense AI — app.py

Main Streamlit application. Wires together extractor.py (Gemini
extraction), analyzer.py (cross-document intelligence) and
excel_builder.py (5-sheet Excel report) behind a 5-step UI:

  1. Upload zone
  2. Document type assignment
  3. Analyse button
  4. Per-document progress
  5. Results tabs (summary / extracted data / issues / download)
"""

from datetime import datetime

import pandas as pd
import streamlit as st

from analyzer import get_doc_amount, run_cross_document_analysis
from excel_builder import build_excel_report
from extractor import (
    DEFAULT_MODEL_NAME,
    DOC_TYPES,
    configure_api,
    extract_document,
    get_all_field_names,
    sanitize_field_name,
)

st.set_page_config(page_title="DocuSense AI", page_icon="🧠", layout="wide")

st.markdown(
    """
    <style>
    div.stButton > button[kind="primary"] {
        height: 3em;
        font-size: 1.1rem;
        font-weight: 600;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Session state
# ---------------------------------------------------------------------------

_DEFAULTS = {
    "doc_type_assignments": {},
    "processing_results": None,
    "analysis_result": None,
    "excel_bytes": None,
    "excel_error": None,
    "processing_stats": None,
    "excluded_fields": set(),
    "custom_fields": [],
}
for _key, _value in _DEFAULTS.items():
    st.session_state.setdefault(_key, _value)


def file_key(file) -> str:
    return f"{file.name}_{file.size}"


def guess_doc_type(filename: str) -> str:
    name = filename.lower()
    if "invoice" in name:
        return "Invoice"
    if "contract" in name or "agreement" in name:
        return "Contract"
    if "po" in name or "purchase" in name or "order" in name:
        return "Purchase Order"
    if "report" in name:
        return "Report"
    return "Other"


# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------

st.markdown(
    """
    <div style="padding: 0.5rem 0 0.2rem 0;">
        <h1 style="margin-bottom:0;">🧠 DocuSense AI</h1>
        <p style="font-size:1.05rem; color:#5a6b7b; margin-top:0.1rem;">
            Upload documents. Get intelligence.
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------

with st.sidebar:
    st.header("⚙️ Configuration")
    api_key = st.text_input(
        "Gemini API Key",
        type="password",
        help="Your key is used only for this session and is never stored or logged.",
    )
    st.markdown("[Get a free Gemini API key →](https://aistudio.google.com/app/apikey)")

    model_name = st.text_input(
        "Gemini Model",
        value=DEFAULT_MODEL_NAME,
        help="Google periodically retires older model versions. If you get a 404/'not found' error, "
        "run genai.list_models() with your key and paste a current flash-tier model name here.",
    )

    st.divider()
    st.subheader("📊 Processing Stats")
    stats = st.session_state.get("processing_stats")
    if stats:
        st.metric("Files Processed", stats["total"])
        c1, c2 = st.columns(2)
        c1.metric("Succeeded", stats["succeeded"])
        c2.metric("Failed", stats["failed"])
        st.metric("Processing Time", f"{stats['duration']:.1f}s")
    else:
        st.caption("Run an analysis to see stats here.")

# ---------------------------------------------------------------------------
# Step 1 — Upload zone
# ---------------------------------------------------------------------------

st.subheader("Step 1 — Upload Documents")
uploaded_files = st.file_uploader(
    "Drag and drop documents here — PDF, PNG, JPG, JPEG, TIF, TIFF (multiple files supported)",
    type=["pdf", "png", "jpg", "jpeg", "tif", "tiff"],
    accept_multiple_files=True,
)

# ---------------------------------------------------------------------------
# Step 2 — Document type assignment
# ---------------------------------------------------------------------------

if uploaded_files:
    st.subheader("Step 2 — Assign Document Types")
    for file in uploaded_files:
        key = file_key(file)
        if key not in st.session_state.doc_type_assignments:
            st.session_state.doc_type_assignments[key] = guess_doc_type(file.name)

        col_name, col_type, col_size = st.columns([3, 2, 1])
        col_name.markdown(f"📄 **{file.name}**")
        selected = col_type.selectbox(
            "Document type",
            DOC_TYPES,
            index=DOC_TYPES.index(st.session_state.doc_type_assignments[key]),
            key=f"type_select_{key}",
            label_visibility="collapsed",
        )
        st.session_state.doc_type_assignments[key] = selected
        col_size.caption(f"{file.size / 1024:.1f} KB")

# ---------------------------------------------------------------------------
# Customize extraction fields (optional)
# ---------------------------------------------------------------------------

st.subheader("🛠️ Customize Extraction Fields (Optional)")
with st.expander("Remove fields you don't need, or add fields of your own"):
    all_field_names = get_all_field_names()

    st.markdown("**Remove fields** — these won't be requested from Gemini for any document")
    st.session_state.excluded_fields = set(
        st.multiselect(
            "Fields to skip",
            options=all_field_names,
            default=sorted(st.session_state.excluded_fields & set(all_field_names)),
            label_visibility="collapsed",
        )
    )
    if st.session_state.excluded_fields & {"po_reference", "po_number", "invoice_number", "expiry_date"}:
        st.caption(
            "⚠️ You've removed a field used for cross-document matching — PO↔invoice matching, "
            "missing-document detection, or contract-expiry checks may be incomplete for affected documents."
        )

    st.markdown("**Add a custom field**")
    cf_cols = st.columns([2, 3, 2, 1, 1])
    new_name = cf_cols[0].text_input("Field name", key="new_field_name", placeholder="e.g. department", label_visibility="collapsed")
    new_desc = cf_cols[1].text_input("Description", key="new_field_desc", placeholder="What should Gemini extract?", label_visibility="collapsed")
    new_scope = cf_cols[2].selectbox("Applies to", ["All Types"] + DOC_TYPES, key="new_field_scope", label_visibility="collapsed")
    new_is_list = cf_cols[3].checkbox("List", key="new_field_is_list", help="Extract as a list of values instead of a single value")

    if cf_cols[4].button("➕ Add"):
        clean_name = sanitize_field_name(new_name)
        existing_names = {cf["name"] for cf in st.session_state.custom_fields}
        if not clean_name:
            st.warning("Enter a field name first.")
        elif clean_name in all_field_names or clean_name in existing_names:
            st.warning(f"'{clean_name}' already exists as a field.")
        else:
            st.session_state.custom_fields.append({
                "name": clean_name,
                "description": new_desc.strip() or f"Custom field: {clean_name}",
                "scope": "All" if new_scope == "All Types" else new_scope,
                "is_list": new_is_list,
            })
            st.rerun()

    if st.session_state.custom_fields:
        st.markdown("**Custom fields added:**")
        remove_idx = None
        for i, cf in enumerate(st.session_state.custom_fields):
            row = st.columns([2, 4, 2, 1])
            row[0].code(cf["name"])
            row[1].caption(cf["description"] + (" (list)" if cf["is_list"] else ""))
            row[2].caption(cf["scope"])
            if row[3].button("🗑️", key=f"remove_cf_{i}"):
                remove_idx = i
        if remove_idx is not None:
            st.session_state.custom_fields.pop(remove_idx)
            st.rerun()

# ---------------------------------------------------------------------------
# Step 3 — Analyse button
# ---------------------------------------------------------------------------

st.subheader("Step 3 — Run Analysis")
button_disabled = not uploaded_files or not api_key
analyse_clicked = st.button(
    "🔍 Analyse Documents", type="primary", width="stretch", disabled=button_disabled
)
if not uploaded_files:
    st.caption("Upload at least one document to continue.")
elif not api_key:
    st.caption("Enter your Gemini API key in the sidebar to continue.")

# ---------------------------------------------------------------------------
# Step 4 — Processing
# ---------------------------------------------------------------------------

if analyse_clicked:
    start_time = datetime.now()

    try:
        model = configure_api(api_key, model_name)
        model.generate_content("Reply with the single word: OK")
    except Exception as e:
        st.error(f"Could not connect to Gemini with this API key: {e}")
        st.stop()

    results = []
    total = len(uploaded_files)
    progress = st.progress(0.0, text="Starting...")

    for i, file in enumerate(uploaded_files):
        doc_type = st.session_state.doc_type_assignments.get(file_key(file), "Other")
        progress.progress(i / total, text=f"Extracting {file.name} ({i + 1}/{total})...")
        try:
            file_bytes = file.getvalue()
            result = extract_document(
                model, file.name, file_bytes, doc_type,
                custom_fields=st.session_state.custom_fields,
                excluded_fields=st.session_state.excluded_fields,
            )
        except Exception as e:
            result = {"filename": file.name, "doc_type": doc_type, "success": False, "data": None, "error": str(e)}
        results.append(result)

    progress.progress(1.0, text="Running cross-document intelligence...")
    try:
        analysis = run_cross_document_analysis(model, results)
    except Exception as e:
        st.warning(f"Cross-document analysis failed: {e}. Document-level results are still available below.")
        analysis = {"matches": [], "issues": [], "vendor_spend": [], "ai_enrichment_success": False, "ai_error": str(e)}

    try:
        custom_field_names = [cf["name"] for cf in st.session_state.custom_fields]
        excel_buffer = build_excel_report(results, analysis, custom_field_names=custom_field_names)
        excel_bytes = excel_buffer.getvalue()
        excel_error = None
    except Exception as e:
        excel_bytes = None
        excel_error = str(e)

    duration = (datetime.now() - start_time).total_seconds()
    succeeded = sum(1 for r in results if r["success"])
    failed = total - succeeded

    st.session_state.processing_results = results
    st.session_state.analysis_result = analysis
    st.session_state.excel_bytes = excel_bytes
    st.session_state.excel_error = excel_error
    st.session_state.processing_stats = {
        "total": total,
        "succeeded": succeeded,
        "failed": failed,
        "duration": duration,
    }

    progress.empty()
    if failed:
        st.warning(f"Processed {total} file(s): {succeeded} succeeded, {failed} failed. See the 'Extracted Data' tab for details.")
    else:
        st.success(f"Successfully processed all {total} document(s).")

# ---------------------------------------------------------------------------
# Step 5 — Results
# ---------------------------------------------------------------------------

results = st.session_state.processing_results
analysis = st.session_state.analysis_result

if results is not None and analysis is not None:
    st.subheader("Step 5 — Results")
    tab1, tab2, tab3, tab4 = st.tabs(
        ["📊 Summary", "📁 Extracted Data", "⚠️ Issues & Flags", "⬇️ Download Excel"]
    )

    with tab1:
        ok_docs = [r for r in results if r["success"]]
        amounts = [get_doc_amount(r["doc_type"], r["data"]) or 0.0 for r in ok_docs]
        currencies = {(r["data"].get("currency") or "").strip() for r in ok_docs if r["data"].get("currency")}
        total_value = sum(amounts)
        if len(currencies) == 1:
            value_label = f"{total_value:,.2f} {next(iter(currencies))}"
        elif currencies:
            value_label = f"{total_value:,.2f} (mixed currencies)"
        else:
            value_label = f"{total_value:,.2f}"

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Documents Processed", len(results))
        c2.metric("Successful Extractions", len(ok_docs))
        c3.metric("Total Value", value_label)
        c4.metric("Issues Found", len(analysis.get("issues", [])))

        if not analysis.get("ai_enrichment_success"):
            st.info(
                f"AI qualitative enrichment skipped ({analysis.get('ai_error')}), but all rule-based "
                "cross-document checks (matching, mismatches, duplicates, expiry, vendor spend) ran normally."
            )

        vendor_spend = analysis.get("vendor_spend", [])
        if vendor_spend:
            st.markdown("**Spend by Vendor**")
            df_vendor = pd.DataFrame(vendor_spend)
            chart_df = df_vendor.set_index("vendor")[["invoice_total", "po_total"]]
            st.bar_chart(chart_df)
            st.dataframe(df_vendor, width="stretch", hide_index=True)

    with tab2:
        for r in results:
            icon = "✅" if r["success"] else "❌"
            with st.expander(f"{icon} {r['filename']} — {r['doc_type']}"):
                if r["success"]:
                    st.json(r["data"])
                else:
                    st.error(r["error"])

    with tab3:
        issues = analysis.get("issues", [])
        if issues:
            df_issues = pd.DataFrame(issues)
            df_issues = df_issues.rename(columns={
                "severity": "Severity",
                "issue_type": "Issue Type",
                "documents_affected": "Document(s) Affected",
                "details": "Details",
                "recommended_action": "Recommended Action",
            })

            def _highlight(row):
                colors = {"High": "#FFCDD2", "Medium": "#FFE0B2", "Low": "#FFF9C4"}
                return [f"background-color: {colors.get(row['Severity'], '')}"] * len(row)

            st.dataframe(df_issues.style.apply(_highlight, axis=1), width="stretch", hide_index=True)
        else:
            st.success("No issues detected across the processed documents.")

    with tab4:
        if st.session_state.excel_bytes:
            st.markdown(
                "Your Excel report includes **5 sheets**: Summary Dashboard, All Documents, "
                "Line Items, Cross-Doc Matches, and Issues & Flags."
            )
            st.download_button(
                "⬇️ Download Excel Report",
                data=st.session_state.excel_bytes,
                file_name=f"DocuSense_AI_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch",
            )
        else:
            st.error(f"Excel report could not be generated: {st.session_state.excel_error}")
