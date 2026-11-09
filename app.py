import os
import uuid
import base64
import streamlit as st
import pandas as pd
from dotenv import load_dotenv

from reconciler import run_reconciliation
from db_audit import (
    get_audit_history_df,
    get_audit_summary_stats,
    record_human_review,
    retrieve_rag_memory,
    log_reconciliation_decision,
)
import re
import importlib
import data_loader
import report_generator
importlib.reload(data_loader)
importlib.reload(report_generator)

from data_loader import (
    load_file_to_df,
    normalize_transactions_df,
    normalize_vendors_df,
)
from report_generator import (
    categorize_transactions_for_report,
    generate_docx_report,
    generate_pdf_report,
    parse_date_sort_key,
    generate_review_queue_csv,
    generate_review_queue_excel,
    generate_review_queue_docx,
    generate_review_queue_pdf,
)

load_dotenv()


def render_csv_download(df: pd.DataFrame, filename: str, button_label: str, button_id: str = "download_btn"):
    """
    Renders an HTML5 data URI download link styled to match Streamlit's primary red button.
    Includes hover and active color transitions like '▶ Run AI Reconciliation'.
    """
    csv_bytes = df.to_csv(index=False).encode("utf-8")
    b64 = base64.b64encode(csv_bytes).decode("utf-8")
    html_code = f"""
    <div style="margin-top: 4px; margin-bottom: 4px;">
        <a id="{button_id}" class="st-primary-download-btn" href="data:text/csv;charset=utf-8;base64,{b64}" download="{filename}">
            {button_label}
        </a>
    </div>
    """
    st.markdown(html_code, unsafe_allow_html=True)

st.set_page_config(
    page_title="AI Reconciliation Agent",
    layout="wide",
    page_icon="🤖"
)

# Custom CSS for polished, modern interface matching Streamlit theme
st.markdown("""
<style>
    .st-primary-download-btn {
        display: block;
        width: 100%;
        text-align: center;
        background-color: #ff4b4b !important;
        color: #ffffff !important;
        font-family: inherit;
        font-weight: 600;
        font-size: 0.95rem;
        line-height: 1.6;
        padding: 0.55rem 1rem;
        border-radius: 0.5rem;
        border: 1px solid #ff4b4b;
        text-decoration: none !important;
        box-shadow: 0 1px 3px rgba(0,0,0,0.12);
        transition: background-color 0.2s ease, border-color 0.2s ease, box-shadow 0.2s ease, transform 0.1s ease;
        box-sizing: border-box;
        cursor: pointer;
    }
    .st-primary-download-btn:hover {
        background-color: #d93838 !important;
        border-color: #d93838 !important;
        color: #ffffff !important;
        box-shadow: 0 3px 10px rgba(255, 75, 75, 0.4);
        transform: translateY(-1px);
    }
    .st-primary-download-btn:active {
        background-color: #b82e2e !important;
        border-color: #b82e2e !important;
        color: #ffffff !important;
        transform: translateY(0);
    }
    .metric-card {
        background-color: #f8f9fa;
        border-radius: 8px;
        padding: 14px 18px;
        border-left: 5px solid #2b6cb0;
        margin-bottom: 10px;
    }
    .badge-human {
        background-color: #e2e8f0;
        color: #2d3748;
        padding: 3px 8px;
        border-radius: 6px;
        font-size: 0.85em;
        font-weight: bold;
    }
    .badge-rag {
        background-color: #ebf8ff;
        color: #2b6cb0;
        padding: 3px 8px;
        border-radius: 6px;
        font-size: 0.85em;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

st.title("🤖 AI Reconciliation Agent")
st.caption(
    f"Powered by AWS Bedrock · **{os.getenv('BEDROCK_MODEL_ID', 'amazon.nova-pro-v1:0')}** · "
    f"Embeddings: **{os.getenv('BEDROCK_EMBEDDING_MODEL_ID', 'amazon.titan-embed-text-v2:0')}** · "
    f"Region: `{os.getenv('AWS_REGION', 'us-east-1')}`"
)

# Top-level Tabs
tab_reconcile, tab_audit, tab_memory = st.tabs([
    "⚡ Reconcile Engine",
    "📜 Traceable Audit Log",
    "🧠 RAG Memory Bank"
])

# ==============================================================================
# TAB 1: RECONCILE ENGINE
# ==============================================================================
with tab_reconcile:
    # Sidebar options
    with st.sidebar:
        st.header("⚙️ Engine Settings")
        use_rag = st.checkbox("🧠 Enable RAG Memory from past reconciled transactions", value=True)
        use_sample = st.checkbox("📂 Use bundled sample data (`sample_data/`)", value=True)
        
        st.subheader("📁 Upload Files")
        st.caption("Supports **CSV**, **PDF** statements/reports, and **Excel** (.xlsx)")
        
        file_1 = st.file_uploader(
            "Bank Transactions (CSV, PDF, Excel)",
            type=["csv", "pdf", "xlsx", "xls", "tsv", "txt"],
            key="txn_file_uploader",
            help="Upload bank/card statements (PDF or CSV) with transaction dates, memos, and amounts."
        )
        file_2 = st.file_uploader(
            "Known Vendors / Chart of Accounts (CSV, PDF, Excel)",
            type=["csv", "pdf", "xlsx", "xls", "tsv", "txt"],
            key="vendor_file_uploader",
            help="Upload vendor lists, expense registers, or chart of accounts (CSV, PDF, or Excel)."
        )
        
        # Option to swap file roles in case user uploaded in reverse order
        swap_files = st.checkbox(
            "⇄ Swap File Roles (File 1 = Vendors, File 2 = Transactions)",
            value=False,
            help="Check this if you uploaded your expense/vendor register in the first box and your bank statement in the second box."
        )
        
        st.divider()
        st.markdown("### 📊 Confidence Thresholds")
        st.markdown("- 🟢 **≥ 0.85** → `AUTO_APPROVE`\n- 🟡 **0.50 – 0.84** → `REVIEW`\n- 🔴 **< 0.50** → `MANUAL`")
        
        st.divider()
        st.markdown(f"**AWS Region:** `{os.getenv('AWS_REGION', 'us-east-1')}`")
        st.markdown(f"**Model:** `{os.getenv('BEDROCK_MODEL_ID', 'amazon.nova-pro-v1:0')}`")

    # Determine which file object is assigned to what based on swap toggle
    if swap_files:
        txn_upload_obj = file_2
        vendor_upload_obj = file_1
    else:
        txn_upload_obj = file_1
        vendor_upload_obj = file_2

    # Data Source Resolution
    txns_df = None
    vendors_df = None
    txn_source_label = "None"
    vendor_source_label = "None"
    txn_mapping_info = {}
    vendor_mapping_info = {}

    sample_txn_path = os.path.join(os.path.dirname(__file__), "sample_data", "bank_transactions.csv")
    sample_vendor_path = os.path.join(os.path.dirname(__file__), "sample_data", "known_vendors.csv")

    # Resolve Transactions (custom upload takes priority over sample)
    if txn_upload_obj is not None:
        raw_t_df, t_info, t_err = load_file_to_df(txn_upload_obj)
        if t_err:
            st.sidebar.error(f"❌ Error reading Transactions ({getattr(txn_upload_obj, 'name', '')}): {t_err}")
        else:
            txns_df, txn_mapping_info = normalize_transactions_df(raw_t_df)
            txn_source_label = f"📁 Uploaded: `{getattr(txn_upload_obj, 'name', '')}` ({t_info} · {len(txns_df)} transactions)"
            st.sidebar.success(f"Transactions: `{getattr(txn_upload_obj, 'name', '')}` ({len(txns_df)} rows loaded)")
    elif use_sample and os.path.exists(sample_txn_path):
        raw_t_df, _, _ = load_file_to_df(sample_txn_path)
        txns_df, txn_mapping_info = normalize_transactions_df(raw_t_df)
        txn_source_label = f"📦 Bundled Sample (`sample_data/`) ({len(txns_df)} rows)"

    # Resolve Vendors / Chart of Accounts (custom upload takes priority over sample)
    if vendor_upload_obj is not None:
        raw_v_df, v_info, v_err = load_file_to_df(vendor_upload_obj)
        if v_err:
            st.sidebar.error(f"❌ Error reading Vendors ({getattr(vendor_upload_obj, 'name', '')}): {v_err}")
        else:
            vendors_df, vendor_mapping_info = normalize_vendors_df(raw_v_df)
            vendor_source_label = f"📁 Uploaded: `{getattr(vendor_upload_obj, 'name', '')}` ({v_info} · {len(vendors_df)} accounts/vendors)"
            st.sidebar.success(f"Vendors: `{getattr(vendor_upload_obj, 'name', '')}` ({len(vendors_df)} accounts loaded)")
    elif use_sample and os.path.exists(sample_vendor_path):
        raw_v_df, _, _ = load_file_to_df(sample_vendor_path)
        vendors_df, vendor_mapping_info = normalize_vendors_df(raw_v_df)
        vendor_source_label = f"📦 Bundled Sample (`sample_data/`) ({len(vendors_df)} rows)"

    # Data preview expanders (auto-expanded when data is loaded)
    has_data = (txns_df is not None and not txns_df.empty) or (vendors_df is not None and not vendors_df.empty)
    with st.expander("👀 Preview Input Data", expanded=has_data):
        col_prev1, col_prev2 = st.columns(2)
        with col_prev1:
            st.markdown(f"**Bank Transactions** — {txn_source_label}")
            if txns_df is not None and not txns_df.empty:
                preview_cols = [c for c in ["id", "date", "memo", "amount"] if c in txns_df.columns]
                other_cols = [c for c in txns_df.columns if c not in preview_cols]
                st.dataframe(txns_df[preview_cols + other_cols], width="stretch", height=200)
                if txn_mapping_info:
                    mapping_str = " · ".join([f"**{k.title()}**: `{v}`" for k, v in txn_mapping_info.items()])
                    st.caption(f"🔍 Column Mapping: {mapping_str}")
            else:
                st.info("No transaction data loaded. Upload a CSV/PDF or check sample data in the sidebar.")
        with col_prev2:
            st.markdown(f"**Known Vendors / Chart of Accounts** — {vendor_source_label}")
            if vendors_df is not None and not vendors_df.empty:
                v_preview_cols = [c for c in ["vendor_name", "gl_account", "gl_description"] if c in vendors_df.columns]
                st.dataframe(vendors_df[v_preview_cols], width="stretch", height=200)
                if vendor_mapping_info:
                    v_mapping_str = " · ".join([f"**{k}**: `{v}`" for k, v in vendor_mapping_info.items()])
                    st.caption(f"🔍 Mapping: {v_mapping_str}")
            else:
                st.info("No vendor data loaded. Upload a CSV/PDF or check sample data in the sidebar.")

    # Action buttons
    run_col, clear_col = st.columns([3, 1])
    with run_col:
        run_btn = st.button("▶ Run AI Reconciliation", type="primary", use_container_width=True)
    with clear_col:
        if st.button("🔄 Reset Results", use_container_width=True):
            st.session_state.pop("df", None)
            st.session_state.pop("current_batch_id", None)
            st.rerun()

    if run_btn:
        if txns_df is None or vendors_df is None or txns_df.empty or vendors_df.empty:
            st.error("Please provide both Bank Transactions and Known Vendors files (CSV, PDF, or Excel).")
        else:
            batch_id = f"batch_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:6]}"
            st.session_state["current_batch_id"] = batch_id
            
            progress_bar = st.progress(0.0)
            status_text = st.empty()
            status_text.text("Nova Pro is analyzing transactions & querying RAG memory...")

            def on_progress(pct, text):
                progress_bar.progress(pct)
                status_text.text(text)

            try:
                results_df = run_reconciliation(
                    txns_df,
                    vendors_df,
                    batch_id=batch_id,
                    use_rag=use_rag,
                    progress_callback=on_progress
                )
                st.session_state["df"] = results_df
                st.success(f"Processed {len(results_df)} transactions! Decisions recorded to SQLite audit log.")
            except Exception as e:
                st.error(f"Reconciliation error: {e}")
            finally:
                progress_bar.empty()
                status_text.empty()

    # Results Display
    if "df" in st.session_state:
        df = st.session_state["df"]

        st.divider()
        st.subheader("📊 Reconciliation Summary")
        
        total_txns = len(df)
        auto_cnt = int((df["status"] == "AUTO_APPROVE").sum())
        review_cnt = int((df["status"] == "REVIEW").sum())
        manual_cnt = int((df["status"] == "MANUAL").sum())
        rag_used_cnt = int((df["rag_memory"] != "None").sum()) if "rag_memory" in df.columns else 0
        
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Total Transactions", total_txns)
        c2.metric("🟢 Auto-Approved", f"{auto_cnt} ({round(auto_cnt/total_txns*100 if total_txns else 0)}%)")
        c3.metric("🟡 Needs Review", f"{review_cnt} ({round(review_cnt/total_txns*100 if total_txns else 0)}%)")
        c4.metric("🔴 Manual Entry", f"{manual_cnt} ({round(manual_cnt/total_txns*100 if total_txns else 0)}%)")
        c5.metric("🧠 RAG Memory Precedents", f"{rag_used_cnt}")

        # Filter controls
        filter_choice = st.radio(
            "Filter Transactions by Status:",
            ["All", "AUTO_APPROVE", "REVIEW", "MANUAL"],
            horizontal=True
        )
        
        if filter_choice != "All":
            display_df = df[df["status"] == filter_choice]
        else:
            display_df = df

        def color_status(val):
            colors = {
                "AUTO_APPROVE": "background-color: #d4edda; color: #155724; font-weight: bold;",
                "REVIEW": "background-color: #fff3cd; color: #856404; font-weight: bold;",
                "MANUAL": "background-color: #f8d7da; color: #721c24; font-weight: bold;",
                "HUMAN_OVERRIDE": "background-color: #cce5ff; color: #004085; font-weight: bold;",
                "HUMAN_APPROVED": "background-color: #d4edda; color: #155724; font-weight: bold;",
            }
            return colors.get(val, "")

        display_cols = [
            "id", "date", "memo", "amount",
            "ai_vendor", "ai_gl_account", "ai_gl_description",
            "confidence", "status", "rag_memory", "reasoning"
        ]
        existing_cols = [c for c in display_cols if c in display_df.columns]
        
        st.dataframe(
            display_df[existing_cols].style.map(color_status, subset=["status"] if "status" in existing_cols else None),
            use_container_width=True,
            height=340,
        )

        # Export Section
        st.subheader("⬇ Export & Reporting")
        
        tab_notes_export, tab_csv_export = st.tabs([
            "📄 Reconciliation Notes (.docx / .pdf)",
            "📊 Data Tables (.csv)"
        ])
        
        with tab_notes_export:
            st.markdown("Generate formatted Reconciliation Notes matching your accounting review format (with bold red section headings, right-aligned amounts, and red highlighted return items).")
            
            # Default title based on transaction dates or current date
            default_title = "September 6 – Notes"
            if "date" in df.columns and not df["date"].empty:
                first_date = str(df["date"].iloc[0]).strip()
                try:
                    if "/" in first_date:
                        parts = first_date.split("/")
                        month_names = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
                        m_idx = int(parts[0]) - 1
                        if 0 <= m_idx < 12:
                            day_val = parts[1].lstrip("0")
                            default_title = f"{month_names[m_idx]} {day_val} – Notes"
                except Exception:
                    pass
            
            c_doc1, c_doc2, c_doc3 = st.columns([3, 3, 2])
            with c_doc1:
                doc_title = st.text_input("Document Title:", value=default_title, key="report_doc_title")
            with c_doc2:
                ledger_heading = st.text_input("Ledger Section Name:", value="Recorded in Appfolio", key="report_ledger_heading")
            with c_doc3:
                max_desc_chars = st.number_input("Max Description Length:", min_value=15, max_value=100, value=30, step=5, key="report_max_desc_len", help="Limits description to 30 characters so rows stay on a single line.")
            
            # Ensure transactions are sorted chronologically ascending by date (August -> September -> October)
            sorted_export_df = df.copy()
            if "date" in sorted_export_df.columns:
                sorted_export_df["_sort_key"] = sorted_export_df["date"].apply(parse_date_sort_key)
                sorted_export_df = sorted_export_df.sort_values(by=["_sort_key"]).drop(columns=["_sort_key"]).reset_index(drop=True)

            # Categorize the records
            report_sections = categorize_transactions_for_report(sorted_export_df, ledger_heading=ledger_heading)
            
            # Show preview summary of sections
            sec_summary_pills = [f"**{sec}**: {len(items)} items" for sec, items in report_sections.items()]
            st.caption("📑 Report Sections: " + " · ".join(sec_summary_pills))
            
            # Generate docx and pdf binaries with 30 character limit
            docx_buffer = generate_docx_report(report_sections, title=doc_title, max_desc_len=int(max_desc_chars))
            pdf_buffer = generate_pdf_report(report_sections, title=doc_title, max_desc_len=int(max_desc_chars))
            
            safe_filename = re.sub(r'[^a-zA-Z0-9_\-]', '_', doc_title.lower()).strip('_')
            
            btn_col1, btn_col2 = st.columns(2)
            with btn_col1:
                st.download_button(
                    label="📄 Download Word Document (.docx)",
                    data=docx_buffer.getvalue(),
                    file_name=f"{safe_filename}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    type="primary"
                )
            with btn_col2:
                st.download_button(
                    label="📑 Download PDF Report (.pdf)",
                    data=pdf_buffer.getvalue(),
                    file_name=f"{safe_filename}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                    type="secondary"
                )
        
        with tab_csv_export:
            col_exp1, col_exp2 = st.columns(2)
            with col_exp1:
                render_csv_download(
                    df,
                    "reconciled_transactions_full.csv",
                    "📥 Download Full Reconciled CSV",
                    button_id="btn_download_full_csv"
                )
            with col_exp2:
                approved_df = df[df["status"].isin(["AUTO_APPROVE", "HUMAN_APPROVED"])]
                render_csv_download(
                    approved_df,
                    "reconciled_approved_entries.csv",
                    "📥 Download Approved Ledger Entries",
                    button_id="btn_download_approved_csv"
                )


        # Human Review Queue
        st.divider()
        st.subheader("👀 Human Review Queue & Active Learning")
        st.caption("Editing below automatically creates traceable human audit records and embeds them into **RAG Memory** for future AI runs.")
        
        review_mask = df["status"].isin(["REVIEW", "MANUAL"])
        review_df = df[review_mask].copy()

        if review_df.empty:
            st.success("🎉 All transactions have been approved!")
        else:
            editable_cols = [
                "id", "date", "memo", "amount",
                "ai_vendor", "ai_gl_account", "ai_gl_description",
                "status", "confidence", "reasoning"
            ]
            review_slice = review_df[[c for c in editable_cols if c in review_df.columns]]
            
            edited = st.data_editor(
                review_slice,
                use_container_width=True,
                num_rows="fixed",
                key="review_editor",
                column_config={
                    "id": st.column_config.TextColumn("ID", disabled=True),
                    "date": st.column_config.TextColumn("Date", disabled=True),
                    "memo": st.column_config.TextColumn("Memo", disabled=True),
                    "amount": st.column_config.NumberColumn("Amount", format="$%.2f", disabled=True),
                    "ai_vendor": st.column_config.TextColumn("Vendor Name"),
                    "ai_gl_account": st.column_config.TextColumn("GL Account"),
                    "ai_gl_description": st.column_config.TextColumn("GL Description"),
                    "status": st.column_config.SelectboxColumn(
                        "Status",
                        options=["AUTO_APPROVE", "REVIEW", "MANUAL"],
                        required=True,
                    ),
                    "reasoning": st.column_config.TextColumn("Accountant Notes / Reasoning"),
                },
            )

            if st.button("✅ Apply Reviews & Learn into RAG Memory", type="primary"):
                current_batch = st.session_state.get("current_batch_id")
                updated_count = 0
                for _, row in edited.iterrows():
                    target_id = row["id"]
                    mask = df["id"] == target_id
                    
                    # Update DataFrame in session
                    for col in edited.columns:
                        if col in df.columns:
                            df.loc[mask, col] = row[col]
                    
                    # Record traceable human audit event and save into RAG Memory
                    new_status = "HUMAN_OVERRIDE" if row["status"] != "MANUAL" else "MANUAL"
                    record_human_review(
                        txn_id=str(target_id),
                        final_vendor=str(row.get("ai_vendor") or ""),
                        final_gl_account=str(row.get("ai_gl_account") or ""),
                        final_gl_description=str(row.get("ai_gl_description") or ""),
                        new_status=new_status,
                        reasoning=str(row.get("reasoning") or "Accountant confirmed/overrode"),
                        batch_id=current_batch
                    )
                    updated_count += 1

                st.session_state["df"] = df
                st.success(f"Applied {updated_count} reviews! Human precedents are now permanently stored in SQLite RAG Memory.")
                st.rerun()

            # Export Review Queue toolbar
            st.markdown("<div style='margin-top: 14px; margin-bottom: 8px; font-weight: 600; font-size: 14px;'>📤 Export Review Queue:</div>", unsafe_allow_html=True)
            
            rq_csv_bytes = generate_review_queue_csv(edited)
            rq_excel_buf = generate_review_queue_excel(edited, title="Human Review Queue & Active Learning Items")
            rq_docx_buf = generate_review_queue_docx(edited, title="Human Review Queue & Active Learning Items", max_desc_len=30)
            rq_pdf_buf = generate_review_queue_pdf(edited, title="Human Review Queue & Active Learning Items", max_desc_len=30)
            
            exp_c1, exp_c2, exp_c3, exp_c4 = st.columns(4)
            with exp_c1:
                st.download_button(
                    label="📥 Export CSV (.csv)",
                    data=rq_csv_bytes,
                    file_name="human_review_queue.csv",
                    mime="text/csv",
                    use_container_width=True,
                    key="btn_export_rq_csv",
                )
            with exp_c2:
                st.download_button(
                    label="📊 Export Excel (.xlsx)",
                    data=rq_excel_buf.getvalue(),
                    file_name="human_review_queue.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="btn_export_rq_excel",
                )
            with exp_c3:
                st.download_button(
                    label="📄 Export Word (.docx)",
                    data=rq_docx_buf.getvalue(),
                    file_name="human_review_queue.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    key="btn_export_rq_docx",
                )
            with exp_c4:
                st.download_button(
                    label="📑 Export PDF (.pdf)",
                    data=rq_pdf_buf.getvalue(),
                    file_name="human_review_queue.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                    key="btn_export_rq_pdf",
                )


# ==============================================================================
# TAB 2: TRACEABLE AUDIT LOG
# ==============================================================================
with tab_audit:
    st.subheader("📜 SQLite Reconciliation Audit Trail")
    st.caption("Every AI match, prompt reasoning, RAG memory citation, and human accountant review is permanently recorded.")

    stats = get_audit_summary_stats()
    s1, s2, s3, s4, s5 = st.columns(5)
    s1.metric("Total Decisions Logged", stats["total_records"])
    s2.metric("🟢 AI Auto-Approved", stats["auto_approved"])
    s3.metric("👤 Human Overrides Learned", stats["human_overrides"])
    s4.metric("⏳ Pending Review", stats["pending_review"])
    s5.metric("🧠 Active RAG Precedents", stats["memory_precedents"])

    st.divider()

    # Search & Filter Controls
    col_f1, col_f2, col_f3 = st.columns([2, 2, 1])
    with col_f1:
        search_kw = st.text_input("🔍 Search Audit Records (memo, vendor, GL, ID):", placeholder="e.g. Starbucks, CHECK, 6100...")
    with col_f2:
        status_filter = st.selectbox(
            "Filter by Decision Status:",
            ["All", "AUTO_APPROVE", "HUMAN_OVERRIDE", "REVIEW", "MANUAL"]
        )
    with col_f3:
        audit_limit = st.selectbox("Record Limit:", [50, 100, 200, 500], index=1)

    audit_df = get_audit_history_df(limit=audit_limit, status_filter=status_filter, search_query=search_kw)

    if audit_df.empty:
        st.info("No audit records found matching your filters. Run a reconciliation batch to generate logs.")
    else:
        st.dataframe(
            audit_df[[
                "id", "created_at", "source", "txn_id", "memo", "amount",
                "vendor_name", "gl_account", "gl_description", "confidence", "status"
            ]],
            use_container_width=True,
            height=320,
        )

        # Traceability Deep Dive Inspector
        st.subheader("🔍 Deep-Dive Decision Inspector")
        record_ids = audit_df["id"].tolist()
        selected_record_id = st.selectbox("Select Audit Record to Inspect Traceability:", record_ids)
        
        if selected_record_id:
            record_row = audit_df[audit_df["id"] == selected_record_id].iloc[0]
            
            with st.container():
                st.markdown(f"### Record #{record_row['id']} — Transaction `{record_row['txn_id']}`")
                c_d1, c_d2, c_d3 = st.columns(3)
                c_d1.markdown(f"**Bank Memo:** `{record_row['memo']}`")
                c_d2.markdown(f"**Amount:** `${record_row['amount']:.2f}`")
                c_d3.markdown(f"**Timestamp:** `{record_row['created_at']}`")
                
                c_d4, c_d5, c_d6 = st.columns(3)
                c_d4.markdown(f"**Decided Vendor:** `{record_row['vendor_name']}`")
                c_d5.markdown(f"**GL Account:** `{record_row['gl_account']}` ({record_row['gl_description']})")
                c_d6.markdown(f"**Decision Source:** `{record_row['source']}` | **Status:** `{record_row['status']}`")

                st.markdown("**🧠 AI Reasoning & Audit Trace:**")
                st.info(record_row["reasoning"])

                # RAG Sources Used
                rag_raw = record_row.get("rag_sources_used")
                if rag_raw and pd.notna(rag_raw):
                    try:
                        import json
                        rag_items = json.loads(rag_raw) if isinstance(rag_raw, str) else rag_raw
                        if rag_items:
                            st.markdown("**📚 RAG Memory Precedents Injected into Prompt:**")
                            for item in rag_items:
                                st.markdown(
                                    f"- **Past Memo:** \"{item.get('memo')}\" ➔ **Vendor:** `{item.get('vendor_name')}` "
                                    f"(GL {item.get('gl_account')}) | Similarity: `{item.get('similarity')}` | Status: `{item.get('status')}`"
                                )
                    except Exception:
                        st.write("RAG metadata format: raw string")

        st.divider()
        render_csv_download(
            audit_df,
            "reconciliation_audit_trail.csv",
            "📥 Download Full Audit Log (CSV)",
            button_id="btn_download_audit_csv"
        )


# ==============================================================================
# TAB 3: RAG MEMORY BANK & SIMULATOR
# ==============================================================================
with tab_memory:
    st.subheader("🧠 RAG Memory Bank & Precedent Simulator")
    st.caption("Test how past approved transactions guide Amazon Nova Pro in reconciling new or ambiguous memos.")

    st.markdown("#### 🧪 Test RAG Retrieval for Any Memo")
    sim_col1, sim_col2 = st.columns([3, 1])
    with sim_col1:
        test_memo = st.text_input("Enter a sample bank memo:", value="SQ *STARBUCKS #9876", placeholder="e.g. AMZN Mktp, SQ *STARBUCKS, CHECK #405...")
    with sim_col2:
        top_k = st.number_input("Top Precedents:", min_value=1, max_value=5, value=3)

    if st.button("🔍 Query RAG Memory Precedents"):
        if not test_memo.strip():
            st.warning("Please enter a memo to query.")
        else:
            with st.spinner("Embedding query & searching memory bank..."):
                retrieved = retrieve_rag_memory(test_memo, top_k=top_k, min_similarity=0.35)
                if not retrieved:
                    st.info(f"No precedents found in memory bank for '{test_memo}' with similarity >= 0.35.")
                else:
                    st.success(f"Found {len(retrieved)} relevant precedent(s) in RAG Memory:")
                    for idx, prec in enumerate(retrieved, 1):
                        st.markdown(f"""
                        **Precedent #{idx} — Similarity Score: `{prec['similarity']}`**
                        - **Past Memo:** `{prec['memo']}`
                        - **Matched Vendor:** `{prec['vendor_name']}`
                        - **GL Account:** `{prec['gl_account']}` ({prec['gl_description']})
                        - **Origin:** `{prec['source']}` ({prec['status']})
                        - **Reasoning:** {prec['reasoning']}
                        """)
                        st.divider()

    st.markdown("#### ➕ Add Permanent Precedent Rule to RAG Memory")
    with st.expander("Register Custom Past Precedent / Vendor Rule", expanded=False):
        with st.form("add_memory_form"):
            c_m1, c_m2 = st.columns(2)
            with c_m1:
                manual_memo = st.text_input("Memo Pattern:", placeholder="e.g. CHECK #405")
                manual_vendor = st.text_input("Vendor Name:", placeholder="e.g. Acme Roofing LLC")
            with c_m2:
                manual_gl = st.text_input("GL Account Code:", placeholder="e.g. 6150")
                manual_desc = st.text_input("GL Description:", placeholder="e.g. Roofing Repairs & Maintenance")
            manual_reason = st.text_area("Accountant Note / Rule Reason:", placeholder="Routine check for roof repair on 5th Ave property.")
            submit_rule = st.form_submit_button("💾 Save Precedent into RAG Memory")

            if submit_rule:
                if not manual_memo or not manual_vendor or not manual_gl:
                    st.error("Please fill in Memo, Vendor, and GL Account.")
                else:
                    log_reconciliation_decision(
                        batch_id="manual_rule",
                        txn_id=f"rule_{uuid.uuid4().hex[:6]}",
                        txn_date=pd.Timestamp.now().strftime("%Y-%m-%d"),
                        memo=manual_memo,
                        amount=0.0,
                        ai_vendor=manual_vendor,
                        ai_gl_account=manual_gl,
                        ai_gl_description=manual_desc,
                        confidence=1.0,
                        status="HUMAN_APPROVED",
                        reasoning=manual_reason or "Accountant defined permanent precedent",
                        model_id="human_rule",
                        source="HUMAN",
                        generate_embed=True
                    )
                    st.success(f"Saved precedent for '{manual_memo}' into RAG Memory! The AI will now use this rule automatically.")