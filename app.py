"""
DocuSense AI — app.py

Main Streamlit application. Wires together extractor.py (Gemini
extraction) and excel_builder.py (3-sheet Excel report: Summary
Dashboard, Header Details, Line Items) behind a 5-step UI:

  1. Upload zone
  2. Document type assignment
  3. Analyse button
  4. Per-document progress
  5. Results tabs (summary / extracted data / download)

Cross-document intelligence (analyzer.py) is intentionally not wired in
here — header + line-item extraction is the only output surfaced. The
module is left intact so it can be reconnected later without a rebuild.
"""

from datetime import datetime

import streamlit as st

from analyzer import normalize_amount
from excel_builder import build_excel_report
from extractor import (
    DEFAULT_MODEL_NAME,
    DOC_TYPES,
    KNOWN_MODELS,
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
    "excel_bytes": None,
    "excel_error": None,
    "processing_stats": None,
    "excluded_fields": set(),
    "custom_fields": [],
    "required_fields": set(),
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

    model_choice = st.selectbox(
        "Gemini Model",
        KNOWN_MODELS + ["Other (type manually)"],
        index=0,
        help="If one model's free-tier daily quota runs out, extraction automatically falls back "
        "through the other models in this list before giving up — each has its own separate quota.",
    )
    if model_choice == "Other (type manually)":
        model_name = st.text_input(
            "Custom model name",
            value=DEFAULT_MODEL_NAME,
            label_visibility="collapsed",
            help="Use this if your key has access to a model not in the dropdown above.",
        )
    else:
        model_name = model_choice

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

    st.markdown("**Mark fields as required** — flagged as a validation issue if missing, alongside any "
                 "field Gemini itself reports low confidence in, and a Subtotal+Tax+Shipping vs Total mismatch")
    requirable_field_names = [f for f in all_field_names if f not in st.session_state.excluded_fields] + [
        cf["name"] for cf in st.session_state.custom_fields
    ]
    st.session_state.required_fields = set(
        st.multiselect(
            "Required fields",
            options=requirable_field_names,
            default=sorted(st.session_state.required_fields & set(requirable_field_names)),
            label_visibility="collapsed",
        )
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
    except Exception as e:
        st.error(f"Could not configure Gemini with this API key: {e}")
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
                required_fields=st.session_state.required_fields,
                primary_model_name=model_name,
            )
        except Exception as e:
            result = {"filename": file.name, "doc_type": doc_type, "success": False, "data": None, "error": str(e)}
        results.append(result)

    progress.progress(1.0, text="Building Excel report...")
    try:
        custom_field_names = [cf["name"] for cf in st.session_state.custom_fields]
        excel_buffer = build_excel_report(results, custom_field_names=custom_field_names)
        excel_bytes = excel_buffer.getvalue()
        excel_error = None
    except Exception as e:
        excel_bytes = None
        excel_error = str(e)

    duration = (datetime.now() - start_time).total_seconds()
    succeeded = sum(1 for r in results if r["success"])
    failed = total - succeeded

    st.session_state.processing_results = results
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

if results is not None:
    st.subheader("Step 5 — Results")
    tab1, tab2, tab3 = st.tabs(["📊 Summary", "📁 Extracted Data", "⬇️ Download Excel"])

    with tab1:
        ok_docs = [r for r in results if r["success"]]
        total_value = sum(normalize_amount(r["data"].get("total")) or 0.0 for r in ok_docs)
        needs_review = sum(1 for r in ok_docs if r.get("validation_issues"))

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Documents Processed", len(results))
        c2.metric("Successful Extractions", len(ok_docs))
        c3.metric("Total Value", f"{total_value:,.2f}")
        c4.metric("Needs Review", needs_review)

    with tab2:
        for r in results:
            icon = "✅" if r["success"] else "❌"
            with st.expander(f"{icon} {r['filename']} — {r['doc_type']}"):
                if r["success"]:
                    model_used = r.get("model_used")
                    if model_used and model_used != model_name:
                        st.caption(f"⚡ Extracted with **{model_used}** — '{model_name}' had hit its daily quota.")
                    issues = r.get("validation_issues") or []
                    if issues:
                        st.warning("**Needs review:**\n" + "\n".join(f"- {issue}" for issue in issues))
                    st.json(r["data"])
                else:
                    st.error(r["error"])

    with tab3:
        if st.session_state.excel_bytes:
            st.markdown(
                "Your Excel report includes **3 sheets**: Summary Dashboard, Header Details, and Line Items."
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
