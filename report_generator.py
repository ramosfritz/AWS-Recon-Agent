import io
import re
from datetime import datetime
from typing import Dict, List, Any, Optional
import pandas as pd

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors


# Common keywords for food, dining, grocery, and general office expenses
FOOD_OFFICE_KEYWORDS = [
    "food", "cafe", "restaurant", "grill", "kabob", "coffee", "starbucks",
    "ralphs", "vons", "wal-mart", "walmart", "cava", "sharky", "panini",
    "grocery", "boudin", "cajas", "sorpresas", "naab", "soomsoom",
    "massis", "office", "supplies", "amazon", "amzn", "mktpl", "parking",
    "laz parking", "prkng", "touchless", "touchless car", "postage", "usps",
    "minute key", "minutekey", "stitch and print", "kimball tirey"
]


def format_amount_str(val: Any) -> str:
    """Formats float/string amount into clean financial string like '1,200.00' or '-10.53'."""
    if pd.isna(val) or val is None:
        return "0.00"
    try:
        f = float(str(val).replace("$", "").replace(",", "").strip())
        return f"{f:,.2f}"
    except (ValueError, TypeError):
        return str(val)


def parse_date_sort_key(d_str: Any) -> tuple:
    """
    Parses date strings into a sortable tuple: (year, month, day).
    Guarantees strict chronological ascending order:
      August (8) -> September (9) -> October (10) -> etc.
    Supports:
      - 'MM/DD/YYYY' / 'M/D/YYYY' / 'MM-DD-YYYY' / 'M-D-YYYY'
      - 'YYYY-MM-DD' / 'YYYY/MM/DD'
      - 'MM/DD' / 'M/D'
      - Strings with timestamps (e.g. '9/29/2026 00:00:00')
      - Blank / invalid dates sort to the end: (9999, 99, 99)
    """
    s = str(d_str).strip()
    if not s or s.lower() in ["none", "nan", "nat", "null"]:
        return (9999, 99, 99)
    
    # Strip any trailing time or whitespace (e.g., '9/29/2026 00:00:00' -> '9/29/2026')
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


def categorize_transactions_for_report(
    df: pd.DataFrame,
    ledger_heading: str = "Recorded in Appfolio"
) -> Dict[str, List[Dict[str, Any]]]:
    """
    Categorizes reconciliation output DataFrame into structured report sections:
    1. Recorded in Appfolio (Approved / ledger matched items)
    2. Returned Item (Refunds, credits, reversing pairs)
    3. Office and Food Expenses (Meals, groceries, dining, office sundries)
    4. Pending Review / Manual (Unmatched or borderline items)
    """
    sections: Dict[str, List[Dict[str, Any]]] = {
        ledger_heading: [],
        "Returned Item": [],
        "Office and Food Expenses": [],
        "Pending Review / Manual": [],
    }

    if df.empty:
        return sections

    records = df.to_dict("records")
    
    # Track pairs for returned items (same absolute amount with +/- or refund keywords)
    amounts_seen: Dict[float, List[int]] = {}

    for idx, r in enumerate(records):
        memo = str(r.get("memo", "")).strip()
        amt_raw = r.get("amount", 0.0)
        try:
            amt = float(str(amt_raw).replace("$", "").replace(",", "").strip())
        except (ValueError, TypeError):
            amt = 0.0

        r["_amt_float"] = amt
        r["_memo_clean"] = memo
        r["_date_clean"] = str(r.get("date", "")).strip()
        r["_formatted_amt"] = format_amount_str(amt)

        abs_amt = round(abs(amt), 2)
        if abs_amt not in amounts_seen:
            amounts_seen[abs_amt] = []
        amounts_seen[abs_amt].append(idx)

    # Detect returned item pairs or standalone return credits
    returned_indices = set()
    for abs_amt, idx_list in amounts_seen.items():
        if abs_amt > 0:
            positives = [i for i in idx_list if records[i]["_amt_float"] > 0]
            negatives = [i for i in idx_list if records[i]["_amt_float"] < 0]
            if positives and negatives:
                returned_indices.update(positives)
                returned_indices.update(negatives)

    for idx, r in enumerate(records):
        memo_lower = r["_memo_clean"].lower()
        amt = r["_amt_float"]
        status = str(r.get("status", "")).upper()
        gl_desc = str(r.get("ai_gl_description", "")).lower()

        # 1. Returned Item check
        if idx in returned_indices or any(k in memo_lower for k in ["return", "refund", "reversal"]):
            sections["Returned Item"].append(r)
            continue

        # 2. Office and Food Expenses check
        is_food_office = any(k in memo_lower for k in FOOD_OFFICE_KEYWORDS) or any(
            k in gl_desc for k in ["meal", "food", "dining", "entertainment", "office", "travel"]
        )
        if is_food_office and status in ["REVIEW", "MANUAL", ""]:
            sections["Office and Food Expenses"].append(r)
            continue

        # 3. Recorded in Ledger (Auto-Approved or Human-Approved)
        if status in ["AUTO_APPROVE", "HUMAN_APPROVED"]:
            sections[ledger_heading].append(r)
            continue

        # 4. Fallback checks
        if is_food_office:
            sections["Office and Food Expenses"].append(r)
        elif status == "REVIEW" or status == "MANUAL":
            sections["Pending Review / Manual"].append(r)
        else:
            sections[ledger_heading].append(r)

    # Clean empty sections and sort each section chronologically by date in ascending order
    cleaned_sections = {}
    for k, v in sections.items():
        if len(v) > 0:
            cleaned_sections[k] = sorted(
                v,
                key=lambda rec: (
                    parse_date_sort_key(rec.get("_date_clean", "")),
                    rec.get("_memo_clean", "").lower(),
                    rec.get("_amt_float", 0.0)
                )
            )
    return cleaned_sections



def generate_docx_report(
    sections: Dict[str, List[Dict[str, Any]]],
    title: str = "Reconciliation Notes",
    max_desc_len: int = 30
) -> io.BytesIO:
    """
    Generates a Microsoft Word document (.docx) with the exact landscape table format
    requested by the user:
    - Landscape orientation with 0.5-inch margins.
    - Soft blue-gray header (#D3DBEB) with bold dark text.
    - 8 columns: Date, ID, Description (Max 30), Amount, Vendor Name, GL Code, Status, Accountant Notes / Reasoning.
    - Sections (e.g. Recorded in Appfolio, Returned Item, Office & Food, Manual) clearly separated with red headers.
    - Chronologically sorted by date ascending.
    - Right-aligned amounts with red highlighting for negative amounts.
    - Alternating subtle zebra row shading.
    """
    doc = Document()

    # Landscape letter orientation
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
        section.page_width = Inches(11.0)
        section.page_height = Inches(8.5)

    # Document Title
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(2)
    run_title = p_title.add_run(title)
    run_title.bold = True
    run_title.font.name = "Arial"
    run_title.font.size = Pt(14)
    run_title.font.color.rgb = RGBColor(30, 41, 59)

    # Subtitle / Summary
    total_records = sum(len(recs) for recs in sections.values())
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(8)
    run_sub = p_sub.add_run(f"Exported: {datetime.now().strftime('%Y-%m-%d %H:%M')} | Total Transactions: {total_records}")
    run_sub.font.name = "Arial"
    run_sub.font.size = Pt(8.5)
    run_sub.font.italic = True
    run_sub.font.color.rgb = RGBColor(100, 116, 139)

    headers = ["Date", "ID", f"Description (Max {max_desc_len})", "Amount", "Vendor Name", "GL Code", "Status", "Accountant Notes / Reasoning"]
    widths = [Inches(0.9), Inches(0.8), Inches(2.2), Inches(0.9), Inches(1.5), Inches(0.8), Inches(0.9), Inches(2.0)]

    red_color = RGBColor(220, 38, 38)
    dark_color = RGBColor(30, 41, 59)

    for section_title, records in sections.items():
        if not records:
            continue

        # Section Header (Red & Bold)
        h_p = doc.add_paragraph()
        h_p.paragraph_format.space_before = Pt(10)
        h_p.paragraph_format.space_after = Pt(4)
        h_run = h_p.add_run(f"{section_title} ({len(records)} items)")
        h_run.bold = True
        h_run.font.name = "Arial"
        h_run.font.size = Pt(11)
        h_run.font.color.rgb = red_color

        # Sort records chronologically ascending by date
        sorted_records = sorted(
            records,
            key=lambda rec: (
                parse_date_sort_key(rec.get("date") or rec.get("_date_clean", "")),
                str(rec.get("memo") or rec.get("_memo_clean", "")).lower(),
                float(rec.get("amount") if rec.get("amount") is not None else rec.get("_amt_float", 0.0))
            )
        )

        tbl = doc.add_table(rows=1, cols=len(headers))
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

        # Format Header Row
        for i, h in enumerate(headers):
            cell = tbl.cell(0, i)
            cell.text = h
            cell.width = widths[i]
            shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="D3DBEB"/>')
            cell._tc.get_or_add_tcPr().append(shd)
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i in [0, 1, 5, 6] else (WD_ALIGN_PARAGRAPH.RIGHT if i == 3 else WD_ALIGN_PARAGRAPH.LEFT)
                for r in p.runs:
                    r.font.name = "Arial"
                    r.font.size = Pt(8.5)
                    r.font.bold = True
                    r.font.color.rgb = dark_color

        # Format Data Rows
        for r_idx, row in enumerate(sorted_records):
            row_cells = tbl.add_row().cells
            raw_memo = str(row.get("memo") or row.get("_memo_clean", "") or "")
            memo_str = raw_memo[:max_desc_len].strip() if max_desc_len and len(raw_memo) > max_desc_len else raw_memo

            amt_raw = row.get("amount") if row.get("amount") is not None else row.get("_amt_float", 0.0)
            try:
                amt_val = float(str(amt_raw).replace("$", "").replace(",", "").strip())
                amt_str = f"{amt_val:,.2f}"
            except Exception:
                amt_val = 0.0
                amt_str = str(amt_raw)

            values = [
                str(row.get("date") or row.get("_date_clean", "") or ""),
                str(row.get("id", "") or ""),
                memo_str,
                amt_str,
                str(row.get("ai_vendor", "") or ""),
                str(row.get("ai_gl_account", "") or ""),
                str(row.get("status", "") or ""),
                str(row.get("reasoning", "") or ""),
            ]

            is_zebra = (r_idx % 2 == 1)
            for i, val in enumerate(values):
                cell = row_cells[i]
                cell.text = val
                cell.width = widths[i]
                if is_zebra:
                    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F8FAFC"/>')
                    cell._tc.get_or_add_tcPr().append(shd)
                for p in cell.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i in [0, 1, 5, 6] else (WD_ALIGN_PARAGRAPH.RIGHT if i == 3 else WD_ALIGN_PARAGRAPH.LEFT)
                    for r in p.runs:
                        r.font.name = "Arial"
                        r.font.size = Pt(8)
                        if i == 3 and amt_val < 0:
                            r.font.color.rgb = red_color
                        else:
                            r.font.color.rgb = dark_color

        p_spacer = doc.add_paragraph()
        p_spacer.paragraph_format.space_before = Pt(3)
        p_spacer.paragraph_format.space_after = Pt(3)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def generate_pdf_report(
    sections: Dict[str, List[Dict[str, Any]]],
    title: str = "Reconciliation Notes",
    max_desc_len: int = 30
) -> io.BytesIO:
    """
    Generates a PDF document report formatted in landscape layout matching the Word table:
    - Landscape letter layout (720pt printable width).
    - Light steel-blue header (#D3DBEB) with dark bold text.
    - 8 columns: Date, ID, Description, Amount, Vendor, GL Code, Status, Notes.
    - Red section headings with counts.
    - Sorted chronologically ascending by date.
    - Right-aligned amounts with negative amounts highlighted in red.
    """
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(letter),
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=36
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "PDFReportTitle",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=14,
        leading=16,
        textColor=colors.HexColor("#1E293B"),
        spaceAfter=2,
    )

    sub_style = ParagraphStyle(
        "PDFReportSub",
        parent=styles["Normal"],
        fontName="Helvetica-Oblique",
        fontSize=8.5,
        leading=10,
        textColor=colors.HexColor("#64748B"),
        spaceAfter=8,
    )

    section_style = ParagraphStyle(
        "PDFSectionHeader",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=11,
        leading=14,
        textColor=colors.HexColor("#DC2626"),
        spaceBefore=10,
        spaceAfter=4,
    )

    cell_style = ParagraphStyle(
        "PDFTableCell",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7.5,
        leading=9,
        textColor=colors.HexColor("#1E293B"),
    )

    total_records = sum(len(recs) for recs in sections.values())
    story = [
        Paragraph(title, title_style),
        Paragraph(f"Exported: {datetime.now().strftime('%Y-%m-%d %H:%M')} | Total Transactions: {total_records}", sub_style),
    ]

    headers = ["Date", "ID", f"Description (Max {max_desc_len})", "Amount", "Vendor Name", "GL Code", "Status", "Accountant Notes / Reasoning"]
    col_widths = [50, 50, 145, 55, 95, 55, 60, 210]

    for section_title, records in sections.items():
        if not records:
            continue

        story.append(Paragraph(f"{section_title} ({len(records)} items)", section_style))

        sorted_records = sorted(
            records,
            key=lambda rec: (
                parse_date_sort_key(rec.get("date") or rec.get("_date_clean", "")),
                str(rec.get("memo") or rec.get("_memo_clean", "")).lower(),
                float(rec.get("amount") if rec.get("amount") is not None else rec.get("_amt_float", 0.0))
            )
        )

        table_data = [headers]
        table_styles = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D3DBEB")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#1E293B")),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 8),
            ("ALIGN", (0, 0), (-1, 0), "LEFT"),
            ("ALIGN", (0, 0), (1, -1), "CENTER"),
            ("ALIGN", (3, 0), (3, -1), "RIGHT"),
            ("ALIGN", (5, 0), (6, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 3.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ]

        for r_idx, row in enumerate(sorted_records):
            raw_memo = str(row.get("memo") or row.get("_memo_clean", "") or "")
            memo_str = raw_memo[:max_desc_len].strip() if max_desc_len and len(raw_memo) > max_desc_len else raw_memo

            amt_raw = row.get("amount") if row.get("amount") is not None else row.get("_amt_float", 0.0)
            try:
                amt_val = float(str(amt_raw).replace("$", "").replace(",", "").strip())
                amt_str = f"{amt_val:,.2f}"
            except Exception:
                amt_val = 0.0
                amt_str = str(amt_raw)

            notes_str = str(row.get("reasoning", "") or "")
            vendor_str = str(row.get("ai_vendor", "") or "")

            row_cells = [
                str(row.get("date") or row.get("_date_clean", "") or ""),
                str(row.get("id", "") or ""),
                Paragraph(memo_str, cell_style),
                amt_str,
                Paragraph(vendor_str, cell_style),
                str(row.get("ai_gl_account", "") or ""),
                str(row.get("status", "") or ""),
                Paragraph(notes_str, cell_style),
            ]
            table_data.append(row_cells)

            data_row_idx = r_idx + 1
            if r_idx % 2 == 1:
                table_styles.append(("BACKGROUND", (0, data_row_idx), (-1, data_row_idx), colors.HexColor("#F8FAFC")))

            if amt_val < 0:
                table_styles.append(("TEXTCOLOR", (3, data_row_idx), (3, data_row_idx), colors.HexColor("#DC2626")))
            else:
                table_styles.append(("TEXTCOLOR", (3, data_row_idx), (3, data_row_idx), colors.HexColor("#1E293B")))

        t = Table(table_data, colWidths=col_widths)
        t.setStyle(TableStyle(table_styles))
        story.append(t)
        story.append(Spacer(1, 6))

    doc.build(story)
    buf.seek(0)
    return buf


def generate_review_queue_csv(df: pd.DataFrame, max_desc_len: Optional[int] = None) -> bytes:
    """
    Exports the review queue to CSV bytes, sorted chronologically ascending by date.
    """
    if df.empty:
        return b""
    df_sorted = df.copy()
    if "date" in df_sorted.columns:
        df_sorted["_sort_key"] = df_sorted["date"].apply(parse_date_sort_key)
        df_sorted = df_sorted.sort_values(by=["_sort_key"]).drop(columns=["_sort_key"]).reset_index(drop=True)
    if max_desc_len and "memo" in df_sorted.columns:
        df_sorted["memo"] = df_sorted["memo"].astype(str).str.slice(0, max_desc_len)

    rename_cols = {
        "id": "ID", "date": "Date", "memo": "Memo", "amount": "Amount",
        "ai_vendor": "Vendor Name", "ai_gl_account": "GL Account",
        "ai_gl_description": "GL Description", "status": "Status",
        "confidence": "Confidence", "reasoning": "Accountant Notes / Reasoning"
    }
    present_cols = [c for c in rename_cols.keys() if c in df_sorted.columns]
    out_df = df_sorted[present_cols].rename(columns={c: rename_cols[c] for c in present_cols})
    return out_df.to_csv(index=False).encode("utf-8")


def generate_review_queue_excel(
    df: pd.DataFrame,
    title: str = "Human Review Queue & Active Learning Items"
) -> io.BytesIO:
    """
    Generates a professionally formatted Excel (.xlsx) file for the review queue:
    - Sorted chronologically ascending by date.
    - Dark navy styled header with white bold text.
    - Currency formatting with negative amounts highlighted in red.
    - Alternating zebra row shading and auto-fitted columns.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Review Queue"
    ws.views.sheetView[0].showGridLines = True

    df_sorted = df.copy()
    if "date" in df_sorted.columns:
        df_sorted["_sort_key"] = df_sorted["date"].apply(parse_date_sort_key)
        df_sorted = df_sorted.sort_values(by=["_sort_key"]).drop(columns=["_sort_key"]).reset_index(drop=True)

    # Title & Metadata
    ws["A1"] = title
    ws["A1"].font = Font(name="Calibri", size=14, bold=True, color="1E293B")
    ws["A2"] = f"Exported: {datetime.now().strftime('%Y-%m-%d %H:%M')} | Total Records: {len(df_sorted)}"
    ws["A2"].font = Font(name="Calibri", size=9, italic=True, color="64748B")

    col_mappings = [
        ("ID", "id"),
        ("Date", "date"),
        ("Memo", "memo"),
        ("Amount", "amount"),
        ("Vendor Name", "ai_vendor"),
        ("GL Account", "ai_gl_account"),
        ("GL Description", "ai_gl_description"),
        ("Status", "status"),
        ("Confidence", "confidence"),
        ("Accountant Notes / Reasoning", "reasoning"),
    ]
    active_cols = [(header, col) for header, col in col_mappings if col in df_sorted.columns]

    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1")
    )
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

    for c_idx, (header_label, _) in enumerate(active_cols, 1):
        cell = ws.cell(row=4, column=c_idx, value=header_label)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = thin_border
        align = "center" if header_label in ["ID", "Date", "GL Account", "Status", "Confidence"] else ("right" if header_label == "Amount" else "left")
        cell.alignment = Alignment(horizontal=align, vertical="center")

    for r_idx, row in df_sorted.iterrows():
        row_num = r_idx + 5
        is_zebra = (r_idx % 2 == 1)
        for c_idx, (_, col_name) in enumerate(active_cols, 1):
            val = row.get(col_name, "")
            cell = ws.cell(row=row_num, column=c_idx)
            cell.border = thin_border
            if is_zebra:
                cell.fill = zebra_fill
            if col_name == "amount":
                try:
                    fval = float(str(val).replace("$", "").replace(",", "").strip())
                    cell.value = fval
                    cell.number_format = "$#,##0.00;($#,##0.00);\"-\""
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    if fval < 0:
                        cell.font = Font(name="Calibri", size=10, color="DC2626")
                    else:
                        cell.font = Font(name="Calibri", size=10, color="1E293B")
                except Exception:
                    cell.value = str(val)
                    cell.font = Font(name="Calibri", size=10)
            elif col_name == "confidence":
                try:
                    fval = float(val)
                    cell.value = fval
                    cell.number_format = "0.0%"
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.font = Font(name="Calibri", size=10)
                except Exception:
                    cell.value = str(val)
                    cell.font = Font(name="Calibri", size=10)
            else:
                cell.value = str(val) if pd.notna(val) else ""
                align = "center" if col_name in ["id", "date", "ai_gl_account", "status"] else "left"
                cell.alignment = Alignment(horizontal=align, vertical="center")
                cell.font = Font(name="Calibri", size=10, color="1E293B")

    for col in ws.columns:
        vals = [str(cell.value or "") for cell in col]
        max_len = max((len(v) for v in vals), default=10)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(min(max_len + 3, 50), 10)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def generate_review_queue_docx(
    df: pd.DataFrame,
    title: str = "Human Review Queue & Active Learning Items",
    max_desc_len: int = 30
) -> io.BytesIO:
    """
    Generates a landscape Word (.docx) document formatted with:
    - Chronological ascending date sort.
    - Description limited to max_desc_len (default: 30 characters).
    - Polished grid table with dark header, clear columns, and red negative amount highlighting.
    """
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
        section.page_width = Inches(11.0)
        section.page_height = Inches(8.5)

    df_sorted = df.copy()
    if "date" in df_sorted.columns:
        df_sorted["_sort_key"] = df_sorted["date"].apply(parse_date_sort_key)
        df_sorted = df_sorted.sort_values(by=["_sort_key"]).drop(columns=["_sort_key"]).reset_index(drop=True)

    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(2)
    r_title = p_title.add_run(title)
    r_title.bold = True
    r_title.font.name = "Arial"
    r_title.font.size = Pt(14)
    r_title.font.color.rgb = RGBColor(30, 41, 59)

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(10)
    r_sub = p_sub.add_run(f"Exported: {datetime.now().strftime('%Y-%m-%d %H:%M')} | Total Items: {len(df_sorted)}")
    r_sub.font.name = "Arial"
    r_sub.font.size = Pt(8.5)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(100, 116, 139)

    headers = ["Date", "ID", f"Description (Max {max_desc_len})", "Amount", "Vendor Name", "GL Code", "Status", "Accountant Notes / Reasoning"]
    widths = [Inches(0.9), Inches(0.8), Inches(2.2), Inches(0.9), Inches(1.5), Inches(0.8), Inches(0.9), Inches(2.0)]

    tbl = doc.add_table(rows=1, cols=len(headers))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

    for i, h in enumerate(headers):
        cell = tbl.cell(0, i)
        cell.text = h
        cell.width = widths[i]
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="D3DBEB"/>')
        cell._tc.get_or_add_tcPr().append(shd)
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i in [0, 1, 5, 6] else (WD_ALIGN_PARAGRAPH.RIGHT if i == 3 else WD_ALIGN_PARAGRAPH.LEFT)
            for r in p.runs:
                r.font.name = "Arial"
                r.font.size = Pt(8.5)
                r.font.bold = True
                r.font.color.rgb = RGBColor(30, 41, 59)

    for r_idx, row in df_sorted.iterrows():
        row_cells = tbl.add_row().cells
        raw_memo = str(row.get("memo", "") or "")
        memo_str = raw_memo[:max_desc_len].strip() if max_desc_len and len(raw_memo) > max_desc_len else raw_memo

        amt_raw = row.get("amount", 0.0)
        try:
            amt_val = float(str(amt_raw).replace("$", "").replace(",", "").strip())
            amt_str = f"{amt_val:,.2f}"
        except Exception:
            amt_val = 0.0
            amt_str = str(amt_raw)

        values = [
            str(row.get("date", "") or ""),
            str(row.get("id", "") or ""),
            memo_str,
            amt_str,
            str(row.get("ai_vendor", "") or ""),
            str(row.get("ai_gl_account", "") or ""),
            str(row.get("status", "") or ""),
            str(row.get("reasoning", "") or ""),
        ]

        is_zebra = (r_idx % 2 == 1)
        for i, val in enumerate(values):
            cell = row_cells[i]
            cell.text = val
            cell.width = widths[i]
            if is_zebra:
                shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F8FAFC"/>')
                cell._tc.get_or_add_tcPr().append(shd)
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i in [0, 1, 5, 6] else (WD_ALIGN_PARAGRAPH.RIGHT if i == 3 else WD_ALIGN_PARAGRAPH.LEFT)
                for r in p.runs:
                    r.font.name = "Arial"
                    r.font.size = Pt(8)
                    if i == 3 and amt_val < 0:
                        r.font.color.rgb = RGBColor(220, 38, 38)
                    else:
                        r.font.color.rgb = RGBColor(30, 41, 59)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def generate_review_queue_pdf(
    df: pd.DataFrame,
    title: str = "Human Review Queue & Active Learning Items",
    max_desc_len: int = 30
) -> io.BytesIO:
    """
    Generates a landscape PDF document for the review queue:
    - Sorted chronologically ascending by date.
    - Description limited to max_desc_len (default: 30 chars).
    - Landscape letter layout fitting all 8 columns comfortably.
    - Auto-wrapping table cells with negative amounts highlighted in red.
    """
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(letter),
        leftMargin=36,
        rightMargin=36,
        topMargin=36,
        bottomMargin=36
    )

    df_sorted = df.copy()
    if "date" in df_sorted.columns:
        df_sorted["_sort_key"] = df_sorted["date"].apply(parse_date_sort_key)
        df_sorted = df_sorted.sort_values(by=["_sort_key"]).drop(columns=["_sort_key"]).reset_index(drop=True)

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "PDFTitle", parent=styles["Normal"], fontName="Helvetica-Bold",
        fontSize=14, leading=16, textColor=colors.HexColor("#1E293B"), spaceAfter=3
    )
    sub_style = ParagraphStyle(
        "PDFSub", parent=styles["Normal"], fontName="Helvetica-Oblique",
        fontSize=8.5, leading=10, textColor=colors.HexColor("#64748B"), spaceAfter=10
    )
    cell_style = ParagraphStyle(
        "PDFCell", parent=styles["Normal"], fontName="Helvetica",
        fontSize=7.5, leading=9, textColor=colors.HexColor("#1E293B")
    )

    story = [
        Paragraph(title, title_style),
        Paragraph(f"Exported: {datetime.now().strftime('%Y-%m-%d %H:%M')} | Total Items: {len(df_sorted)}", sub_style),
    ]

    headers = ["Date", "ID", f"Description (Max {max_desc_len})", "Amount", "Vendor Name", "GL Code", "Status", "Accountant Notes / Reasoning"]
    table_data = [headers]
    col_widths = [50, 50, 145, 55, 95, 55, 60, 210]

    table_styles = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D3DBEB")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#1E293B")),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 8),
        ("ALIGN", (0, 0), (-1, 0), "LEFT"),
        ("ALIGN", (0, 0), (1, -1), "CENTER"),
        ("ALIGN", (3, 0), (3, -1), "RIGHT"),
        ("ALIGN", (5, 0), (6, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
    ]

    for r_idx, row in df_sorted.iterrows():
        raw_memo = str(row.get("memo", "") or "")
        memo_str = raw_memo[:max_desc_len].strip() if max_desc_len and len(raw_memo) > max_desc_len else raw_memo

        amt_raw = row.get("amount", 0.0)
        try:
            amt_val = float(str(amt_raw).replace("$", "").replace(",", "").strip())
            amt_str = f"{amt_val:,.2f}"
        except Exception:
            amt_val = 0.0
            amt_str = str(amt_raw)

        notes_str = str(row.get("reasoning", "") or "")
        vendor_str = str(row.get("ai_vendor", "") or "")

        row_cells = [
            str(row.get("date", "") or ""),
            str(row.get("id", "") or ""),
            Paragraph(memo_str, cell_style),
            amt_str,
            Paragraph(vendor_str, cell_style),
            str(row.get("ai_gl_account", "") or ""),
            str(row.get("status", "") or ""),
            Paragraph(notes_str, cell_style),
        ]
        table_data.append(row_cells)

        data_row_idx = r_idx + 1
        if r_idx % 2 == 1:
            table_styles.append(("BACKGROUND", (0, data_row_idx), (-1, data_row_idx), colors.HexColor("#F8FAFC")))

        if amt_val < 0:
            table_styles.append(("TEXTCOLOR", (3, data_row_idx), (3, data_row_idx), colors.HexColor("#DC2626")))
        else:
            table_styles.append(("TEXTCOLOR", (3, data_row_idx), (3, data_row_idx), colors.HexColor("#1E293B")))

    t = Table(table_data, colWidths=col_widths)
    t.setStyle(TableStyle(table_styles))
    story.append(t)

    doc.build(story)
    buf.seek(0)
    return buf
