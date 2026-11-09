# 🤖 AWS Bedrock AI Reconciliation Agent

An intelligent, CSV-based financial transaction reconciliation application powered by **AWS Bedrock (Amazon Nova Pro)** and **Streamlit**.

---

## 🚀 Features

- **Amazon Nova Pro Reasoning**: Utilizes Bedrock Converse API with tool use (`submit_match`) to accurately perform fuzzy matching between bank transaction memos and vendor GL accounts.
- **SQLite Audit Trail (Traceable AI Decisions)**: Every reconciliation decision, confidence score, model ID, reasoning, and accountant review is logged in SQLite (`recon_agent.db`) with full search, filtering, and CSV export.
- **Continuous RAG Memory (Active Learning)**:
  - Powered by **Amazon Titan Embeddings v2** (256-dim normalized vectors) + lexical fuzzy matching.
  - Automatically retrieves relevant past approved and human-overridden decisions to guide Nova Pro on similar transactions.
  - Human accountant edits in the Review Queue immediately become high-priority precedents in RAG memory.
- **Three-Tier Confidence Categorization**:
  - 🟢 **AUTO_APPROVE** ($\ge 0.85$ confidence): High confidence matches ready for ledger sync.
  - 🟡 **REVIEW** ($0.50 - 0.84$ confidence): Borderline matches requiring accountant confirmation.
  - 🔴 **MANUAL** ($< 0.50$ confidence or unmatched): No high-confidence match found (e.g. generic checks, unknown memos).
- **Interactive Review Queue & Active Learning**: Edit vendor names, GL codes, and statuses directly in Streamlit and persist edits back to the dataset and memory bank.
- **Dual Export Options**: Download full reconciled results or only the auto-approved ledger transactions.
- **Bundled Sample Data**: Includes realistic bank transaction memos and chart of accounts in `sample_data/` for instant one-click testing.

---

## 🛠️ Setup & Installation

### 1. Prerequisites
- Python 3.10+
- AWS IAM credentials with Amazon Bedrock permissions (`amazon.nova-pro-v1:0` access enabled in your AWS region)

### 2. Environment Configuration
Ensure your [`.env`](file:///c:/Users/DELL/Documents/aws_recon_agent/.env) file has:

```env
# AWS Bedrock Configuration
AWS_REGION=us-east-1
AWS_ACCESS_KEY_ID=your_access_key_here
AWS_SECRET_ACCESS_KEY=your_secret_access_key_here
BEDROCK_MODEL_ID=amazon.nova-pro-v1:0
```

### 3. Virtual Environment & Dependencies
```powershell
# Create virtual environment (if not already created)
python -m venv venv

# Activate virtual environment
# Windows:
.\venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

---

## 🖥️ Running the Application

Start the Streamlit dashboard:

```powershell
streamlit run app.py
```

The app will be available at:
`http://localhost:8501`

---

## 📁 Supported Input File Formats (CSV, PDF, Excel)

The application automatically supports multiple file formats and intelligent column normalization:

### Supported File Types
- **PDF Statements**: Bank and credit card statements (e.g. Chase, Wells Fargo, BofA, etc.) are automatically parsed into transactions with dates, memos, and amounts.
- **CSV & Delimited Files**: Automatic multi-encoding detection (`utf-8`, `latin1`, `cp1252`, `iso-8859-1`) and automatic separator sniffing (comma, semicolon, tab).
- **Excel Spreadsheets**: `.xlsx` and `.xls` files.

### 1. Bank Transactions (CSV or PDF)
- Auto-maps columns matching:
  - **Memo / Description**: `memo`, `description`, `narrative`, `payee name`, `payee`, `merchant`, `details`
  - **Amount**: `amount`, `debit`, `credit`, `total` (cleans currency symbols `$`, commas, and parenthesized negatives)
  - **Date**: `date`, `bill date`, `check date`, `posting date`, `transaction date`
  - **ID**: `id`, `reference/check #`, `check number`, `doc #` (or auto-generates IDs)

### 2. Known Vendors / Chart of Accounts / Expense Register (CSV or PDF)
- Supports standard tables (`vendor_name`, `gl_account`, `gl_description`).
- Supports accounting Expense Registers with `Expense Account` (e.g., `6130 - Maintenance: Flooring`) and `Payee Name` / `Description`.

