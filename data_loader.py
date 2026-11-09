import io
import re
import os
from typing import Tuple, Dict, Any, Optional, List
import pandas as pd
from pypdf import PdfReader


def parse_date_sort_key(d_str: Any) -> tuple:
    """
    Parses date strings into a sortable tuple: (year, month, day).
    Guarantees strict chronological ascending order:
      August (8) -> September (9) -> October (10) -> etc.
    """
    s = str(d_str).strip()
    if not s or s.lower() in ["none", "nan", "nat", "null"]:
        return (9999, 99, 99)
    
    # Strip any trailing time or whitespace (e.g. '9/29/2026 00:00:00')
    s_date = s.split()[0].strip()

    # 1. ISO format: YYYY-MM-DD or YYYY/MM/DD
    m_iso = re.match(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})", s_date)
    if m_iso:
        return (int(m_iso.group(1)), int(m_iso.group(2)), int(m_iso.group(3)))

    # 2. US format with year: M/D/YYYY, MM/DD/YYYY, M/D/YY, MM/DD/YY, M-D-YYYY, etc.
    m_us_yr = re.match(r"^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})", s_date)
    if m_us_yr:
        month = int(m_us_yr.group(1))
        day = int(m_us_yr.group(2))
        yr_val = int(m_us_yr.group(3))
        year = 2000 + yr_val if yr_val < 100 else yr_val
        return (year, month, day)

    # 3. US format without year: MM/DD or M/D or MM-DD or M-D
    m_us_noyr = re.match(r"^(\d{1,2})[-/](\d{1,2})$", s_date)
    if m_us_noyr:
        month = int(m_us_noyr.group(1))
        day = int(m_us_noyr.group(2))
        return (2026, month, day)

    # 4. Fallback with pandas to_datetime
    try:
        dt = pd.to_datetime(s_date, errors="coerce")
        if pd.notna(dt):
            return (int(dt.year), int(dt.month), int(dt.day))
    except Exception:
        pass

    return (9999, 99, 99)


def clean_amount_val(val: Any) -> float:
    """
    Safely converts amount representations into float.
    Handles "$1,200.00", "1,200.00", "(24.22)", "-$50.00", etc.
    """
    if pd.isna(val) or val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    
    s = str(val).strip()
    if not s:
        return 0.0
    
    # Check parenthesized negative amounts: (123.45) -> -123.45
    is_negative = False
    if s.startswith("(") and s.endswith(")"):
        is_negative = True
        s = s[1:-1].strip()
    elif s.startswith("-"):
        is_negative = True
        s = s[1:].strip()
    elif s.endswith("-"):
        is_negative = True
        s = s[:-1].strip()
        
    s = s.replace("$", "").replace(",", "").replace(" ", "").replace("€", "").replace("£", "")
    try:
        amt = float(s)
        return -amt if is_negative else amt
    except ValueError:
        return 0.0


def extract_transactions_from_pdf(pdf_stream_or_path: Any) -> pd.DataFrame:
    """
    Extracts tabular financial transactions from bank/credit card PDF statements.
    Works with Chase, Wells Fargo, BofA, Citi, Amex, and standard statement layouts.
    """
    if isinstance(pdf_stream_or_path, (str, os.PathLike)):
        reader = PdfReader(pdf_stream_or_path)
    else:
        # File-like object (e.g. Streamlit UploadedFile)
        if hasattr(pdf_stream_or_path, "seek"):
            pdf_stream_or_path.seek(0)
        reader = PdfReader(pdf_stream_or_path)

    txns: List[Dict[str, Any]] = []
    
    # Common statement skip patterns (headers, footers, summaries)
    skip_keywords = [
        "page ", "statement date", "customer service", "account activity",
        "payment due date", "new balance", "minimum payment", "late payment",
        "opening/closing", "annual percentage", "interest charged", "total fees",
        "notice:", "telephone:", "p.o. box", "manage your account"
    ]

    # Regex patterns for line-based transaction records
    # Pattern 1: MM/DD or MM/DD/YY[YY] followed by description and amount at the end
    # Example: '08/06 RINGCENTRAL INC. 888-898-4591 CA -24.22' or '08/10 WWW COSTCO COM WA 1,233.36'
    pattern_standard = re.compile(
        r"^(\d{1,2}/\d{1,2}(?:/\d{2,4})?)\s+(.+?)\s+([-\$]?\s*[\d,]+\.\d{2}-?)$"
    )
    
    # Pattern 2: YYYY-MM-DD or DD-MM-YYYY format
    pattern_iso = re.compile(
        r"^(\d{4}[-/]\d{1,2}[-/]\d{1,2}|\d{1,2}[-/]\d{1,2}[-/]\d{4})\s+(.+?)\s+([-\$]?\s*[\d,]+\.\d{2}-?)$"
    )

    txn_idx = 1
    for page_num, page in enumerate(reader.pages, start=1):
        text = page.extract_text() or ""
        lines = text.split("\n")
        
        for line in lines:
            line_str = line.strip()
            if not line_str or len(line_str) < 6:
                continue
            
            lower_line = line_str.lower()
            if any(k in lower_line for k in skip_keywords):
                continue

            match = pattern_standard.match(line_str) or pattern_iso.match(line_str)
            if match:
                date_str, desc, amt_str = match.groups()
                desc_clean = desc.strip()
                amt_clean = clean_amount_val(amt_str)
                
                # Filter out obvious false positives (e.g. account numbers or phone numbers misparsed)
                if len(desc_clean) > 1 and not desc_clean.startswith("XXXX"):
                    txns.append({
                        "id": f"pdf_txn_{txn_idx:04d}",
                        "date": date_str.strip(),
                        "memo": desc_clean,
                        "amount": amt_clean,
                        "page": page_num,
                    })
                    txn_idx += 1

    return pd.DataFrame(txns)


def extract_vendors_from_pdf(pdf_stream_or_path: Any) -> pd.DataFrame:
    """
    Extracts vendor / GL account information from a PDF list or report.
    """
    if isinstance(pdf_stream_or_path, (str, os.PathLike)):
        reader = PdfReader(pdf_stream_or_path)
    else:
        if hasattr(pdf_stream_or_path, "seek"):
            pdf_stream_or_path.seek(0)
        reader = PdfReader(pdf_stream_or_path)

    vendors: List[Dict[str, Any]] = []
    # Pattern: AccountCode - Category Description or Vendor - Account
    pattern_gl = re.compile(r"^(\d{3,6})\s*[-–]\s*(.+)$")
    
    for page in reader.pages:
        text = page.extract_text() or ""
        for line in text.split("\n"):
            line_str = line.strip()
            m = pattern_gl.match(line_str)
            if m:
                code, desc = m.groups()
                vendors.append({
                    "vendor_name": desc.strip(),
                    "gl_account": code.strip(),
                    "gl_description": desc.strip(),
                })

    if not vendors:
        # Fallback: treat transactions extracted as vendor precedents
        txns_df = extract_transactions_from_pdf(pdf_stream_or_path)
        if not txns_df.empty:
            for _, r in txns_df.iterrows():
                vendors.append({
                    "vendor_name": r["memo"][:40],
                    "gl_account": "6000",
                    "gl_description": "General Expense",
                })

    return pd.DataFrame(vendors).drop_duplicates(subset=["vendor_name", "gl_account"])


def load_file_to_df(file_obj_or_path: Any, file_name: Optional[str] = None) -> Tuple[pd.DataFrame, str, str]:
    """
    Robust reader for CSV, PDF, Excel (XLSX/XLS), and TSV files.
    Returns: (DataFrame, detected_format_info, error_message)
    """
    if file_name is None:
        if isinstance(file_obj_or_path, str):
            file_name = os.path.basename(file_obj_or_path)
        elif hasattr(file_obj_or_path, "name"):
            file_name = getattr(file_obj_or_path, "name", "unknown")
        else:
            file_name = "unknown"

    lower_name = file_name.lower()
    
    # 1. Handle PDF
    if lower_name.endswith(".pdf"):
        try:
            df = extract_transactions_from_pdf(file_obj_or_path)
            if not df.empty:
                return df, f"PDF Document ({len(df)} transactions extracted)", ""
            # If transactions were empty, try vendor parser
            v_df = extract_vendors_from_pdf(file_obj_or_path)
            if not v_df.empty:
                return v_df, f"PDF Document ({len(v_df)} accounts/vendors extracted)", ""
            return pd.DataFrame(), "PDF Document (0 records found)", "No tabular records could be extracted from PDF."
        except Exception as e:
            return pd.DataFrame(), "PDF Parse Error", str(e)

    # 2. Handle Excel
    if lower_name.endswith((".xlsx", ".xls")):
        try:
            if hasattr(file_obj_or_path, "seek"):
                file_obj_or_path.seek(0)
            df = pd.read_excel(file_obj_or_path)
            df = df.dropna(how="all")
            return df, f"Excel Spreadsheet ({len(df)} rows)", ""
        except Exception as e:
            return pd.DataFrame(), "Excel Parse Error", str(e)

    # 3. Handle CSV / TSV / TXT
    encodings_to_try = ["utf-8", "utf-8-sig", "latin1", "cp1252", "iso-8859-1", "windows-1252"]
    last_err = ""
    
    for enc in encodings_to_try:
        try:
            if hasattr(file_obj_or_path, "seek"):
                file_obj_or_path.seek(0)
            # Try Python engine with separator detection
            df = pd.read_csv(
                file_obj_or_path,
                encoding=enc,
                sep=None,
                engine="python",
                on_bad_lines="skip"
            )
            df = df.dropna(how="all")
            if not df.empty and len(df.columns) > 1:
                return df, f"CSV/Delimited ({enc}, {len(df)} rows)", ""
        except Exception as e:
            last_err = str(e)
            continue

    # Fallback to standard comma separated if sep=None had trouble
    for enc in ["latin1", "utf-8", "cp1252"]:
        try:
            if hasattr(file_obj_or_path, "seek"):
                file_obj_or_path.seek(0)
            df = pd.read_csv(file_obj_or_path, encoding=enc, on_bad_lines="skip")
            df = df.dropna(how="all")
            if not df.empty:
                return df, f"CSV Standard ({enc}, {len(df)} rows)", ""
        except Exception as e:
            last_err = str(e)
            continue

    return pd.DataFrame(), "Parse Failed", last_err or "Unable to parse file."


def normalize_transactions_df(df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """
    Normalizes arbitrary transaction data into standardized columns:
    ['id', 'date', 'memo', 'amount'].
    Preserves all original columns as well.
    Returns: (normalized_df, mapping_used)
    """
    if df.empty:
        return df, {}

    cols_lower = {str(c).strip().lower(): c for c in df.columns}
    mapping: Dict[str, str] = {}
    
    # 1. Resolve Memo / Description
    memo_candidates = [
        "memo", "description", "narrative", "payee name", "payee",
        "merchant", "details", "transaction description", "name", "trans_desc"
    ]
    matched_memo = None
    for cand in memo_candidates:
        if cand in cols_lower:
            matched_memo = cols_lower[cand]
            break
    if not matched_memo:
        # Fallback to the first string/text column with high uniqueness
        for c in df.columns:
            if df[c].dtype == object and c.lower() not in ["id", "date", "bill date", "check date"]:
                matched_memo = c
                break
    mapping["memo"] = matched_memo or "memo"

    # 2. Resolve Amount
    amt_candidates = [
        "amount", "debit", "credit", "net amount", "total", "value",
        "amount ($)", "amt", "trans_amount"
    ]
    matched_amt = None
    for cand in amt_candidates:
        if cand in cols_lower:
            matched_amt = cols_lower[cand]
            break
    mapping["amount"] = matched_amt or "amount"

    # 3. Resolve Date
    date_candidates = [
        "date", "check date", "bill date", "posting date",
        "transaction date", "post date", "trans date", "dt"
    ]
    matched_date = None
    for cand in date_candidates:
        if cand in cols_lower:
            matched_date = cols_lower[cand]
            break
    mapping["date"] = matched_date or "date"

    # 4. Resolve ID
    id_candidates = [
        "id", "reference/check #", "reference", "check #",
        "check number", "txn_id", "trans id", "doc #", "ref", "num"
    ]
    matched_id = None
    for cand in id_candidates:
        if cand in cols_lower:
            matched_id = cols_lower[cand]
            break
    mapping["id"] = matched_id or "id"

    # Build clean output dataframe
    clean_df = df.copy()
    
    # Clean Amount column
    if matched_amt and matched_amt in clean_df.columns:
        clean_df["amount"] = clean_df[matched_amt].apply(clean_amount_val)
    else:
        clean_df["amount"] = 0.0

    # Clean Memo column
    if matched_memo and matched_memo in clean_df.columns:
        clean_df["memo"] = clean_df[matched_memo].fillna("").astype(str).str.strip()
    else:
        clean_df["memo"] = ""

    # Clean Date column
    if matched_date and matched_date in clean_df.columns:
        clean_df["date"] = clean_df[matched_date].fillna("").astype(str).str.strip()
    else:
        clean_df["date"] = pd.Timestamp.now().strftime("%Y-%m-%d")

    # Clean ID column
    if matched_id and matched_id in clean_df.columns:
        clean_df["id"] = clean_df[matched_id].fillna("").astype(str).str.strip()
        # If any IDs are empty, fill with synthetic ID
        empty_mask = (clean_df["id"] == "") | (clean_df["id"] == "nan")
        if empty_mask.any():
            clean_df.loc[empty_mask, "id"] = [f"txn_{i+1:04d}" for i in range(empty_mask.sum())]
    else:
        clean_df["id"] = [f"txn_{i+1:04d}" for i in range(len(clean_df))]

    # Filter out empty spacer rows (common in accounting software exports)
    valid_mask = (clean_df["memo"] != "") & (clean_df["amount"] != 0.0)
    filtered_df = clean_df[valid_mask].copy()
    
    # If filtering removed everything, keep original non-null rows
    if filtered_df.empty:
        filtered_df = clean_df[clean_df["memo"] != ""].copy()
    if filtered_df.empty:
        filtered_df = clean_df

    # Always sort transactions chronologically by date in ascending order (August before September, etc.)
    if "date" in filtered_df.columns:
        filtered_df["_sort_key"] = filtered_df["date"].apply(parse_date_sort_key)
        filtered_df = filtered_df.sort_values(by=["_sort_key", "memo", "amount"]).drop(columns=["_sort_key"])

    # Reset indices
    filtered_df = filtered_df.reset_index(drop=True)
    return filtered_df, mapping


def normalize_vendors_df(df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """
    Normalizes arbitrary vendor or Chart of Accounts data into standardized columns:
    ['vendor_name', 'gl_account', 'gl_description'].
    Also handles Expense Registers containing 'Expense Account' (e.g. '6130 - Maintenance: Flooring').
    """
    if df.empty:
        return df, {}

    cols_lower = {str(c).strip().lower(): c for c in df.columns}
    mapping: Dict[str, str] = {}
    
    # Case 1: Check if this is an Expense Register with 'Expense Account'
    exp_acc_col = None
    for cand in ["expense account", "expense_account", "gl account", "chart of accounts"]:
        if cand in cols_lower:
            exp_acc_col = cols_lower[cand]
            break

    desc_col = None
    for cand in ["description", "memo", "item description"]:
        if cand in cols_lower:
            desc_col = cols_lower[cand]
            break

    payee_col = None
    for cand in ["payee name", "payee", "vendor", "vendor_name", "merchant", "supplier"]:
        if cand in cols_lower:
            payee_col = cols_lower[cand]
            break

    # If we have an Expense Account column like "6130 - Maintenance: Flooring", parse out vendors & GLs
    if exp_acc_col and (desc_col or payee_col):
        vendor_records: List[Dict[str, str]] = []
        valid_rows = df[df[exp_acc_col].notna() & (df[exp_acc_col] != "")].copy()
        
        for _, r in valid_rows.iterrows():
            exp_acc = str(r.get(exp_acc_col, "")).strip()
            gl_code = ""
            gl_desc = exp_acc
            if " - " in exp_acc:
                parts = exp_acc.split(" - ", 1)
                gl_code = parts[0].strip()
                gl_desc = parts[1].strip()
            else:
                gl_code = exp_acc[:6].strip()

            # Identify vendor name
            v_name = ""
            if desc_col and pd.notna(r.get(desc_col)):
                raw_d = str(r.get(desc_col)).strip()
                if " - " in raw_d:
                    v_name = raw_d.split(" - ", 1)[0].strip()
                else:
                    v_name = raw_d[:40].strip()
            
            if not v_name and payee_col and pd.notna(r.get(payee_col)):
                v_name = str(r.get(payee_col)).strip()

            if v_name and gl_code:
                vendor_records.append({
                    "vendor_name": v_name,
                    "gl_account": gl_code,
                    "gl_description": gl_desc,
                })

        if vendor_records:
            v_df = pd.DataFrame(vendor_records).drop_duplicates(subset=["vendor_name", "gl_account"]).reset_index(drop=True)
            mapping["vendor_name"] = f"Extracted from {desc_col or payee_col}"
            mapping["gl_account"] = f"Parsed from {exp_acc_col}"
            mapping["gl_description"] = f"Parsed from {exp_acc_col}"
            return v_df, mapping

    # Case 2: Standard or near-standard table
    clean_df = df.copy()

    # Vendor Name
    v_cands = ["vendor_name", "vendor", "payee name", "payee", "merchant", "supplier", "company", "name"]
    matched_v = None
    for cand in v_cands:
        if cand in cols_lower:
            matched_v = cols_lower[cand]
            break
    mapping["vendor_name"] = matched_v or "vendor_name"

    # GL Account
    gl_cands = ["gl_account", "expense account", "account", "account #", "gl code", "gl_code", "category", "code"]
    matched_gl = None
    for cand in gl_cands:
        if cand in cols_lower:
            matched_gl = cols_lower[cand]
            break
    mapping["gl_account"] = matched_gl or "gl_account"

    # GL Description
    desc_cands = ["gl_description", "expense description", "account description", "category description", "description"]
    matched_desc = None
    for cand in desc_cands:
        if cand in cols_lower:
            matched_desc = cols_lower[cand]
            break
    mapping["gl_description"] = matched_desc or "gl_description"

    if matched_v and matched_v in clean_df.columns:
        clean_df["vendor_name"] = clean_df[matched_v].fillna("").astype(str).str.strip()
    else:
        clean_df["vendor_name"] = clean_df.iloc[:, 0].astype(str)

    if matched_gl and matched_gl in clean_df.columns:
        clean_df["gl_account"] = clean_df[matched_gl].fillna("").astype(str).str.strip()
    else:
        clean_df["gl_account"] = "6000"

    if matched_desc and matched_desc in clean_df.columns:
        clean_df["gl_description"] = clean_df[matched_desc].fillna("").astype(str).str.strip()
    else:
        clean_df["gl_description"] = clean_df["gl_account"]

    # Filter out empty vendors
    clean_df = clean_df[clean_df["vendor_name"] != ""].copy()
    clean_df = clean_df.drop_duplicates(subset=["vendor_name", "gl_account"]).reset_index(drop=True)
    return clean_df, mapping
