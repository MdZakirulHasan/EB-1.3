"""
DUET Individual Examination Bill Generator
------------------------------------------
Single-page Streamlit application.

Install:
    pip install streamlit pandas openpyxl reportlab

Run:
    streamlit run individual_exam_bill_app.py

Upload the approved examination remuneration-rate CSV in the application.
The app creates one individual teacher/person bill at a time.

IMPORTANT:
- The app does not authenticate users or approve/pay bills.
- Minimum/maximum scopes and sharing rules must be selected/confirmed
  according to the applicable official order.
- It does not infer eligibility or quantities from a teacher's name.
"""

from __future__ import annotations

import io
import re
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Any

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


APP_TITLE = "Individual Examination Bill Generator"
UNIVERSITY = "Dhaka University of Engineering & Technology (DUET), Gazipur"
DEPARTMENTS = [
    "Civil Engineering (CE)",
    "Electrical and Electronic Engineering (EEE)",
    "Mechanical Engineering (ME)",
    "Computer Science and Engineering (CSE)",
    "Textile Engineering (TE)",
    "Industrial & Production Engineering (IPE)",
    "Architecture",
    "Food Engineering (FE)",
    "Chemical Engineering (ChE)",
    "Materials and Metallurgical Engineering (MME)",
]
TEACHING_DEPARTMENTS = DEPARTMENTS + [
    "Chemistry", "Mathematics", "Physics", "Humanities & Social Sciences", "Other"
]
CENT = Decimal("0.01")


# ---------------------------------------------------------------------------
# Utility and CSV handling
# ---------------------------------------------------------------------------

def money(value: Any) -> Decimal:
    """Convert a value to Decimal, safely handling common BDT formatting."""
    if value is None:
        return Decimal("0")
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "null", "***", "-"}:
        return Decimal("0")
    text = text.replace(",", "").replace("BDT", "").replace("Tk.", "").replace("Tk", "").strip()
    text = re.sub(r"[^0-9.\-]", "", text)
    if text in {"", "-", ".", "-."}:
        return Decimal("0")
    try:
        return Decimal(text)
    except InvalidOperation as exc:
        raise ValueError(f"Could not interpret monetary value: {value}") from exc


def dec(value: Any, field: str) -> Decimal:
    try:
        result = Decimal(str(value).strip())
    except (InvalidOperation, ValueError, AttributeError):
        raise ValueError(f"Enter a valid number for {field}.")
    if not result.is_finite() or result < 0:
        raise ValueError(f"{field} must be a non-negative number.")
    return result


def fmt(value: Decimal | float | int) -> str:
    return f"{Decimal(str(value)).quantize(CENT, rounding=ROUND_HALF_UP):,.2f}"


def clean_excel_text(value: Any) -> Any:
    """Prevent user-entered text from being interpreted as an Excel formula."""
    if isinstance(value, str) and value.startswith(("=", "+", "-", "@")):
        return "'" + value
    return value


def normalize_column(name: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(name).strip().lower())


COLUMN_ALIASES = {
    "serial": {"slno", "slno.", "serial", "serialno", "sl"},
    "item": {
        "item", "description", "descriptionofwork", "nameofitem",
        "billitem", "work", "examinationactivity", "particulars"
    },
    "rate": {"rate", "ratebdt", "amountperunit", "unitrate", "remuneration"},
    "unit": {"unit", "unitofpayment", "quantityunit", "basis"},
    "condition": {"condition", "conditions", "remarks", "note", "descriptioncondition"},
    "minimum": {"minimum", "min", "minimumrate", "minimumamount", "minimumbdt"},
    "maximum": {"maximum", "max", "maximumrate", "maximumamount", "maximumbdt"},
}


def find_column(columns: list[Any], key: str) -> Any | None:
    normalized = {normalize_column(c): c for c in columns}
    for alias in COLUMN_ALIASES[key]:
        if alias in normalized:
            return normalized[alias]
    # Partial matching for common extended headings.
    for norm, original in normalized.items():
        if key == "item" and any(x in norm for x in ("descriptionofwork", "examinationactivity", "billitem")):
            return original
        if key == "serial" and norm.startswith("slno"):
            return original
        if key == "minimum" and norm.startswith("minimum"):
            return original
        if key == "maximum" and norm.startswith("maximum"):
            return original
    return None


def read_rate_csv(uploaded_file) -> pd.DataFrame:
    """Read and normalize an uploaded CSV rate schedule."""
    raw = None
    errors = []
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
        try:
            uploaded_file.seek(0)
            raw = pd.read_csv(uploaded_file, sep=None, engine="python", encoding=encoding)
            break
        except Exception as exc:
            errors.append(str(exc))
    if raw is None:
        raise ValueError("Could not read the CSV file. " + " | ".join(errors[-2:]))

    raw.columns = [str(c).strip() for c in raw.columns]
    mapping = {key: find_column(list(raw.columns), key) for key in COLUMN_ALIASES}
    if mapping["item"] is None or mapping["rate"] is None:
        raise ValueError(
            "The CSV must contain at least an item/description column and a rate column. "
            f"Detected headings: {', '.join(raw.columns)}"
        )

    normalized = pd.DataFrame()
    normalized["serial"] = raw[mapping["serial"]].fillna("").astype(str) if mapping["serial"] else [
        str(i + 1) for i in range(len(raw))
    ]
    normalized["item"] = raw[mapping["item"]].fillna("").astype(str).str.strip()
    normalized["rate_text"] = raw[mapping["rate"]].fillna("").astype(str).str.strip()
    normalized["unit"] = raw[mapping["unit"]].fillna("").astype(str).str.strip() if mapping["unit"] else ""
    normalized["condition"] = raw[mapping["condition"]].fillna("").astype(str).str.strip() if mapping["condition"] else ""
    normalized["minimum"] = raw[mapping["minimum"]].apply(money) if mapping["minimum"] else Decimal("0")
    normalized["maximum"] = raw[mapping["maximum"]].apply(money) if mapping["maximum"] else Decimal("0")
    normalized = normalized[normalized["item"].str.len() > 0].reset_index(drop=True)

    if normalized.empty:
        raise ValueError("No examination bill items were found in the CSV.")
    normalized["choice"] = normalized.apply(
        lambda r: f'{r["serial"]}. {r["item"]}', axis=1
    )
    return normalized


def calculate_line(
    quantity: Decimal,
    rate: Decimal,
    sharing: str,
    beneficiaries: int,
    limit_rule: str,
    limit_scope: str,
    minimum: Decimal,
    maximum: Decimal,
) -> dict[str, Decimal | str]:
    """Calculate the individual person's payable amount and show each step."""
    if quantity <= 0:
        raise ValueError("Quantity must be greater than zero.")
    if rate < 0:
        raise ValueError("Rate cannot be negative.")
    count = beneficiaries if sharing == "Equal sharing" else 1
    if count < 1:
        raise ValueError("Number of beneficiaries must be at least one.")

    gross_total = quantity * rate
    individual_before_limits = gross_total / Decimal(count)

    if limit_scope == "Total amount before sharing":
        adjusted_base = gross_total
    else:
        adjusted_base = individual_before_limits

    adjusted = adjusted_base
    if limit_rule in {"Apply minimum", "Apply minimum and maximum"} and minimum > 0:
        adjusted = max(adjusted, minimum)
    if limit_rule in {"Apply maximum", "Apply minimum and maximum"} and maximum > 0:
        adjusted = min(adjusted, maximum)

    if limit_scope == "Total amount before sharing":
        individual_payable = adjusted / Decimal(count)
        total_payable = adjusted
    else:
        individual_payable = adjusted
        total_payable = adjusted * Decimal(count)

    individual_payable = individual_payable.quantize(CENT, rounding=ROUND_HALF_UP)
    total_payable = total_payable.quantize(CENT, rounding=ROUND_HALF_UP)

    return {
        "gross_total": gross_total,
        "individual_before_limits": individual_before_limits,
        "individual_payable": individual_payable,
        "total_payable": total_payable,
        "beneficiaries": Decimal(count),
        "adjustment": limit_rule,
        "limit_scope": limit_scope,
    }


# ---------------------------------------------------------------------------
# Export helpers
# ---------------------------------------------------------------------------

def make_excel(bill: dict, lines: list[dict]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Individual Bill"
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:J1")
    ws["A1"] = UNIVERSITY
    ws["A1"].font = Font(name="Arial", size=14, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center")

    ws.merge_cells("A2:J2")
    ws["A2"] = "INDIVIDUAL EXAMINATION BILL"
    ws["A2"].font = Font(name="Arial", size=13, bold=True)
    ws["A2"].alignment = Alignment(horizontal="center")

    meta = [
        ("Bill reference", bill["reference"]),
        ("Examination", bill["exam_name"]),
        ("Academic session / batch", bill["session"]),
        ("Examination year", bill["exam_year"]),
        ("Semester", bill["semester"]),
        ("Examination category", bill["category"]),
        ("Bill-owning department", bill["bill_department"]),
        ("Teacher / member", bill["teacher"]),
        ("Designation", bill["designation"]),
        ("Home department", bill["home_department"]),
        ("Preparation date", bill["prepared_date"]),
        ("Status", "DRAFT — NOT APPROVED / NOT PAID"),
    ]
    for idx, (label, value) in enumerate(meta, start=4):
        ws.cell(idx, 1, label).font = Font(bold=True)
        ws.cell(idx, 2, clean_excel_text(str(value)))

    header_row = 4 + len(meta) + 1
    headers = [
        "Sl.", "Bill item / duty", "Course / subject", "Teaching department",
        "Class group", "Quantity", "Unit", "Rate (BDT)",
        "Calculation / sharing details", "Payable to this teacher (BDT)"
    ]
    for col, heading in enumerate(headers, 1):
        c = ws.cell(header_row, col, heading)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="17365D")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    thin = Side(style="thin", color="A6A6A6")
    for i, line in enumerate(lines, start=1):
        row = header_row + i
        values = [
            i, line["item"], line["course"], line["teaching_department"],
            line["class_group"], float(line["quantity"]), line["unit"],
            float(line["rate"]), line["calculation_note"],
            float(line["individual_payable"])
        ]
        for col, value in enumerate(values, 1):
            c = ws.cell(row, col, clean_excel_text(value))
            c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            c.alignment = Alignment(vertical="top", wrap_text=True)
        for col in (8, 10):
            ws.cell(row, col).number_format = '#,##0.00'

    total_row = header_row + len(lines) + 1
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=9)
    ws.cell(total_row, 1, "TOTAL PAYABLE AMOUNT (BDT)").font = Font(bold=True)
    ws.cell(total_row, 1).alignment = Alignment(horizontal="right")
    ws.cell(total_row, 10, float(sum((x["individual_payable"] for x in lines), Decimal("0"))))
    ws.cell(total_row, 10).font = Font(bold=True)
    ws.cell(total_row, 10).number_format = '#,##0.00'

    note_row = total_row + 2
    ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=10)
    ws.cell(note_row, 1, "Note: This is a generated draft. Verify eligibility, quantities, rates, sharing and minimum/maximum scopes before approval.")
    ws.cell(note_row, 1).alignment = Alignment(wrap_text=True)
    ws.cell(note_row, 1).font = Font(italic=True, size=9)

    sig_row = note_row + 4
    ws.cell(sig_row, 1, "Prepared by:").font = Font(bold=True)
    ws.cell(sig_row, 4, "Checked by:").font = Font(bold=True)
    ws.cell(sig_row, 7, "Approved by:").font = Font(bold=True)
    ws.cell(sig_row + 4, 1, "Signature of teacher/member:").font = Font(bold=True)
    ws.cell(sig_row + 4, 7, "Signature and seal:").font = Font(bold=True)

    widths = [7, 31, 20, 24, 16, 12, 18, 16, 45, 24]
    for col, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.freeze_panes = f"A{header_row + 1}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


def make_pdf(bill: dict, lines: list[dict]) -> bytes:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    )

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=landscape(A4),
        rightMargin=12*mm, leftMargin=12*mm, topMargin=10*mm, bottomMargin=10*mm
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "BillTitle", parent=styles["Title"], fontName="Helvetica-Bold",
        fontSize=13, leading=16, alignment=TA_CENTER, spaceAfter=3
    )
    small = ParagraphStyle("Small", parent=styles["BodyText"], fontSize=7, leading=9)
    story = [
        Paragraph(UNIVERSITY, title_style),
        Paragraph("INDIVIDUAL EXAMINATION BILL", title_style),
        Spacer(1, 5*mm)
    ]

    meta_rows = [
        ["Bill reference", bill["reference"], "Examination", bill["exam_name"]],
        ["Academic session / batch", bill["session"], "Examination year", bill["exam_year"]],
        ["Semester / category", f'{bill["semester"]} / {bill["category"]}',
         "Bill-owning department", bill["bill_department"]],
        ["Teacher / member", bill["teacher"], "Designation", bill["designation"]],
        ["Home department", bill["home_department"], "Preparation date", bill["prepared_date"]],
    ]
    mt = Table(meta_rows, colWidths=[35*mm, 65*mm, 35*mm, 105*mm])
    mt.setStyle(TableStyle([
        ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
        ("FONTNAME", (0,0), (0,-1), "Helvetica-Bold"),
        ("FONTNAME", (2,0), (2,-1), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 8),
        ("GRID", (0,0), (-1,-1), .35, colors.grey),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("BACKGROUND", (0,0), (0,-1), colors.whitesmoke),
        ("BACKGROUND", (2,0), (2,-1), colors.whitesmoke),
        ("LEFTPADDING", (0,0), (-1,-1), 5),
        ("RIGHTPADDING", (0,0), (-1,-1), 5),
        ("TOPPADDING", (0,0), (-1,-1), 5),
        ("BOTTOMPADDING", (0,0), (-1,-1), 5),
    ]))
    story += [mt, Spacer(1, 5*mm)]

    data = [[
        "Sl.", "Bill item / duty", "Course", "Teaching dept.",
        "Class group", "Qty.", "Unit", "Rate (BDT)",
        "Calculation / sharing", "Payable (BDT)"
    ]]
    for i, line in enumerate(lines, 1):
        data.append([
            str(i),
            Paragraph(str(line["item"]), small),
            Paragraph(str(line["course"]), small),
            Paragraph(str(line["teaching_department"]), small),
            Paragraph(str(line["class_group"]), small),
            str(line["quantity"]),
            Paragraph(str(line["unit"]), small),
            fmt(line["rate"]),
            Paragraph(str(line["calculation_note"]), small),
            fmt(line["individual_payable"]),
        ])
    total = sum((x["individual_payable"] for x in lines), Decimal("0"))
    data.append(["", "", "", "", "", "", "", "", "TOTAL", fmt(total)])

    col_widths = [8*mm, 48*mm, 25*mm, 29*mm, 19*mm, 13*mm, 20*mm, 20*mm, 55*mm, 25*mm]
    table = Table(data, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#17365D")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 7),
        ("GRID", (0,0), (-1,-2), .35, colors.grey),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("ALIGN", (0,0), (0,-1), "CENTER"),
        ("ALIGN", (5,1), (7,-1), "RIGHT"),
        ("ALIGN", (9,1), (9,-1), "RIGHT"),
        ("BACKGROUND", (0,-1), (-1,-1), colors.whitesmoke),
        ("FONTNAME", (8,-1), (9,-1), "Helvetica-Bold"),
        ("SPAN", (0,-1), (7,-1)),
        ("ALIGN", (0,-1), (7,-1), "RIGHT"),
        ("TOPPADDING", (0,0), (-1,-1), 5),
        ("BOTTOMPADDING", (0,0), (-1,-1), 5),
    ]))
    story += [table, Spacer(1, 5*mm)]
    story.append(Paragraph(
        "<b>Draft notice:</b> This generated bill is not an approval or payment order. "
        "Verify eligibility, quantities, rates, sharing and minimum/maximum scopes before approval.",
        small
    ))
    story.append(Spacer(1, 15*mm))
    sig = Table(
        [["________________________", "________________________", "________________________"],
         ["Prepared by", "Checked by", "Approved by"],
         ["", "", ""],
         ["________________________", "________________________", "________________________"],
         ["Teacher/member signature", "Department/committee", "Authorized signature and seal"]],
        colWidths=[85*mm, 85*mm, 85*mm]
    )
    sig.setStyle(TableStyle([
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("FONTSIZE", (0,0), (-1,-1), 8),
        ("FONTNAME", (0,1), (-1,1), "Helvetica-Bold"),
        ("FONTNAME", (0,4), (-1,4), "Helvetica-Bold"),
        ("TOPPADDING", (0,0), (-1,-1), 4),
    ]))
    story.append(sig)
    doc.build(story)
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# Streamlit single-page application
# ---------------------------------------------------------------------------

st.set_page_config(page_title=APP_TITLE, page_icon="🧾", layout="wide")

st.markdown("""
<style>
    .block-container {padding-top: 1.4rem; padding-bottom: 2rem; max-width: 1500px;}
    .app-subtitle {color: #64748b; margin-top: -12px; margin-bottom: 18px;}
    .bill-total {background: #f1f5f9; padding: 14px 18px; border-radius: 10px;
                 border: 1px solid #cbd5e1;}
    .draft-label {color: #9a3412; font-weight: 700; letter-spacing: .04em;}
    div[data-testid="stMetric"] {background: #f8fafc; border: 1px solid #e2e8f0;
                                 padding: 10px; border-radius: 9px;}
</style>
""", unsafe_allow_html=True)

st.title("🧾 Individual Examination Bill Generator")
st.markdown(
    '<div class="app-subtitle">Prepare, review and download an individual teacher / committee-member examination bill.</div>',
    unsafe_allow_html=True
)
st.caption("Dhaka University of Engineering & Technology (DUET), Gazipur")
st.warning(
    "This application generates a DRAFT bill only. It does not verify institutional eligibility, "
    "approve a bill or authorize payment. Confirm all rates and official conditions before use.",
    icon="⚠️"
)

# Keep entries during Streamlit reruns.
if "bill_lines" not in st.session_state:
    st.session_state.bill_lines = []
if "rate_table" not in st.session_state:
    st.session_state.rate_table = None
if "rate_filename" not in st.session_state:
    st.session_state.rate_filename = ""

# Rate schedule
st.subheader("1. Load examination remuneration rates")
rate_col1, rate_col2 = st.columns([2, 1])
with rate_col1:
    uploaded_rates = st.file_uploader(
        "Upload the approved rate CSV",
        type=["csv"],
        help="The CSV should contain item/description and rate columns. Unit, condition, minimum and maximum columns are recommended."
    )
with rate_col2:
    template = pd.DataFrame([
        {"Sl. No.": "1", "Item": "Example: Answer-script examination",
         "Rate": "100", "Unit": "Per answer script per section",
         "Condition": "Minimum Tk. 1,000 per examiner per section",
         "Minimum": "1000", "Maximum": "0"}
    ])
    st.download_button(
        "Download CSV template",
        data=template.to_csv(index=False).encode("utf-8-sig"),
        file_name="exam_bill_rate_template.csv",
        mime="text/csv",
        use_container_width=True
    )

if uploaded_rates is not None:
    try:
        st.session_state.rate_table = read_rate_csv(uploaded_rates)
        st.session_state.rate_filename = uploaded_rates.name
        st.success(f"Rate schedule loaded: {uploaded_rates.name} ({len(st.session_state.rate_table)} items).")
    except Exception as exc:
        st.session_state.rate_table = None
        st.error(f"Could not load the rate CSV: {exc}")

rates_df = st.session_state.rate_table
if rates_df is None:
    st.info("Upload your examination bill-rate CSV to begin. The application will use the rates and conditions in that file.")
    st.stop()

with st.expander("View loaded rate schedule", expanded=False):
    display_rates = rates_df[["serial", "item", "rate_text", "unit", "condition", "minimum", "maximum"]].copy()
    display_rates.columns = ["Sl. No.", "Bill item", "Rate", "Unit", "Condition", "Minimum (BDT)", "Maximum (BDT)"]
    st.dataframe(display_rates, use_container_width=True, hide_index=True)

st.divider()

# Examination details
st.subheader("2. Examination information")
c1, c2, c3 = st.columns(3)
with c1:
    bill_department = st.selectbox("Bill for which department?", DEPARTMENTS)
    exam_name = st.text_input("Examination / programme name", placeholder="e.g., B.Sc. Engineering Examination")
with c2:
    academic_session = st.text_input("Academic session / batch", placeholder="e.g., 2022–2023")
    exam_year = st.text_input("Examination year", value=str(date.today().year))
with c3:
    semester = st.selectbox("Semester", ["1st Semester", "2nd Semester"])
    category = st.selectbox("Examination category", ["Regular", "Review", "Special", "Other"])

c4, c5, c6 = st.columns(3)
with c4:
    class_group = st.text_input("Class group", placeholder="e.g., 3rd Year 1st Semester")
with c5:
    bill_reference = st.text_input("Bill reference (optional)", placeholder="Auto-generated if blank")
with c6:
    prepared_date = st.date_input("Bill preparation date", value=date.today())

st.divider()

# Teacher details
st.subheader("3. Teacher / committee-member information")
t1, t2, t3 = st.columns(3)
with t1:
    teacher_name = st.text_input("Teacher / member name")
    designation = st.text_input("Designation")
with t2:
    home_department = st.selectbox("Teacher's home department", TEACHING_DEPARTMENTS)
    employee_id = st.text_input("Employee ID (optional)")
with t3:
    committee_role = st.text_input("Committee / assigned role", placeholder="e.g., Examiner / Moderator")
    teacher_email = st.text_input("Institutional email (optional)")

st.divider()

# Bill line entry
st.subheader("4. Add examination duties / bill items")
item_options = rates_df["choice"].tolist()
selected_choice = st.selectbox("Select bill item", item_options, key="selected_bill_item")
selected_row = rates_df.loc[rates_df["choice"] == selected_choice].iloc[0]
rate_is_manual = selected_row["rate_text"].strip() == "***"

with st.container(border=True):
    item_col1, item_col2, item_col3 = st.columns([2, 1, 1])
    with item_col1:
        st.markdown(f"**{selected_row['item']}**")
        st.caption(f"Official unit: {selected_row['unit'] or 'Not specified'}")
        if selected_row["condition"]:
            st.info(f"Rate condition: {selected_row['condition']}")
    with item_col2:
        st.metric("Minimum (BDT)", fmt(selected_row["minimum"]))
    with item_col3:
        st.metric("Maximum (BDT)", fmt(selected_row["maximum"]))

    q1, q2, q3 = st.columns(3)
    with q1:
        quantity_text = st.text_input("Quantity", value="1", help="Enter the actual eligible quantity for this duty.")
    with q2:
        if rate_is_manual:
            rate_text_input = st.text_input("Approved manual rate (BDT)", placeholder="Enter authorized rate")
        else:
            rate_text_input = st.text_input("Rate (BDT)", value=str(selected_row["rate_text"]), disabled=True)
    with q3:
        unit_override = st.text_input("Unit (editable)", value=str(selected_row["unit"]))

    st.markdown("**Course / duty details**")
    d1, d2, d3 = st.columns(3)
    with d1:
        course_code_title = st.text_input("Course code / title", placeholder="Leave blank if not course-specific")
    with d2:
        teaching_department = st.selectbox("Course-owning / teaching department", TEACHING_DEPARTMENTS, key="teaching_dept")
    with d3:
        line_class_group = st.text_input("Class group for this item", value=class_group)

    st.markdown("**Sharing and minimum / maximum application**")
    s1, s2, s3 = st.columns(3)
    with s1:
        sharing = st.selectbox("Distribution", ["Individual payment", "Equal sharing"])
    with s2:
        beneficiaries = st.number_input(
            "Total eligible beneficiaries", min_value=1, max_value=100, value=1, step=1,
            disabled=(sharing == "Individual payment")
        )
    with s3:
        limit_rule = st.selectbox(
            "Apply rate limits?",
            ["Do not apply automatically", "Apply minimum", "Apply maximum", "Apply minimum and maximum"],
            help="Select only after confirming that the minimum/maximum applies at the selected scope."
        )

    scope_options = ["This teacher's share", "Total amount before sharing"]
    limit_scope_label = st.selectbox(
        "Minimum / maximum scope",
        scope_options,
        disabled=(limit_rule == "Do not apply automatically"),
        help="Choose the scope stated in the approved remuneration order."
    )

    add_line = st.button("＋ Add this item to the individual bill", type="primary", use_container_width=True)

if add_line:
    try:
        qty = dec(quantity_text, "quantity")
        if qty <= 0:
            raise ValueError("Quantity must be greater than zero.")
        if rate_is_manual:
            rate_value = dec(rate_text_input, "manual rate")
        else:
            rate_value = money(selected_row["rate_text"])
        if rate_value <= 0:
            raise ValueError("The rate must be greater than zero. For a rate shown as ***, enter an approved manual rate.")

        min_value = Decimal(str(selected_row["minimum"]))
        max_value = Decimal(str(selected_row["maximum"]))
        rule_map = {
            "Do not apply automatically": "No automatic adjustment",
            "Apply minimum": "Apply minimum",
            "Apply maximum": "Apply maximum",
            "Apply minimum and maximum": "Apply minimum and maximum",
        }
        scope_map = {
            "This teacher's share": "This teacher's share",
            "Total amount before sharing": "Total amount before sharing",
        }
        result = calculate_line(
            quantity=qty,
            rate=rate_value,
            sharing=sharing,
            beneficiaries=int(beneficiaries),
            limit_rule=rule_map[limit_rule],
            limit_scope=scope_map[limit_scope_label],
            minimum=min_value,
            maximum=max_value,
        )

        if sharing == "Equal sharing":
            share_text = f'Equal sharing among {int(beneficiaries)} eligible persons'
        else:
            share_text = "Individual payment; no sharing"
        calculation_note = (
            f'{qty} × BDT {fmt(rate_value)} = BDT {fmt(result["gross_total"])} total; '
            f'{share_text}; {result["adjustment"]}; scope: {result["limit_scope"]}. '
            f'Payable to this person: BDT {fmt(result["individual_payable"])}.'
        )
        if limit_rule != "Do not apply automatically" and (min_value > 0 or max_value > 0):
            st.session_state.bill_lines.append({
                "item": str(selected_row["item"]),
                "serial": str(selected_row["serial"]),
                "course": course_code_title.strip(),
                "teaching_department": teaching_department,
                "class_group": line_class_group.strip(),
                "quantity": qty,
                "unit": unit_override.strip(),
                "rate": rate_value,
                "condition": str(selected_row["condition"]),
                "minimum": min_value,
                "maximum": max_value,
                "gross_total": result["gross_total"],
                "individual_payable": result["individual_payable"],
                "total_payable": result["total_payable"],
                "sharing": share_text,
                "calculation_note": calculation_note,
                "rate_source": st.session_state.rate_filename,
            })
        else:
            st.session_state.bill_lines.append({
                "item": str(selected_row["item"]),
                "serial": str(selected_row["serial"]),
                "course": course_code_title.strip(),
                "teaching_department": teaching_department,
                "class_group": line_class_group.strip(),
                "quantity": qty,
                "unit": unit_override.strip(),
                "rate": rate_value,
                "condition": str(selected_row["condition"]),
                "minimum": min_value,
                "maximum": max_value,
                "gross_total": result["gross_total"],
                "individual_payable": result["individual_payable"],
                "total_payable": result["total_payable"],
                "sharing": share_text,
                "calculation_note": calculation_note,
                "rate_source": st.session_state.rate_filename,
            })
        st.success("Bill item added.")
    except Exception as exc:
        st.error(str(exc))

# Current bill
st.divider()
st.subheader("5. Individual bill preview")

if st.session_state.bill_lines:
    preview_rows = []
    for idx, line in enumerate(st.session_state.bill_lines, start=1):
        preview_rows.append({
            "Sl.": idx,
            "Bill item": line["item"],
            "Course": line["course"],
            "Quantity": str(line["quantity"]),
            "Unit": line["unit"],
            "Rate (BDT)": fmt(line["rate"]),
            "Sharing": line["sharing"],
            "Payable to teacher (BDT)": fmt(line["individual_payable"]),
        })
    st.dataframe(pd.DataFrame(preview_rows), use_container_width=True, hide_index=True)

    grand_total = sum((x["individual_payable"] for x in st.session_state.bill_lines), Decimal("0"))
    total_col, count_col, remove_col = st.columns([2, 1, 2])
    with total_col:
        st.markdown(
            f'<div class="bill-total"><div class="draft-label">TOTAL PAYABLE — DRAFT</div>'
            f'<h2 style="margin:4px 0 0 0;">BDT {fmt(grand_total)}</h2></div>',
            unsafe_allow_html=True
        )
    with count_col:
        st.metric("Number of bill items", len(st.session_state.bill_lines))
    with remove_col:
        remove_index = st.selectbox(
            "Remove an item",
            options=list(range(len(st.session_state.bill_lines))),
            format_func=lambda i: f'{i + 1}. {st.session_state.bill_lines[i]["item"]}',
            key="remove_line_index"
        )
        if st.button("Remove selected item", use_container_width=True):
            st.session_state.bill_lines.pop(remove_index)
            st.rerun()

    with st.expander("View calculation details and rate conditions"):
        for idx, line in enumerate(st.session_state.bill_lines, 1):
            st.markdown(f"**{idx}. {line['item']}**")
            st.write(line["calculation_note"])
            if line["condition"]:
                st.caption("CSV condition: " + line["condition"])
            st.caption(
                f"Minimum: BDT {fmt(line['minimum'])} | "
                f"Maximum: BDT {fmt(line['maximum'])} | "
                f"Rate source: {line['rate_source']}"
            )

    # Validate required teacher and examination fields before export.
    missing_fields = []
    if not teacher_name.strip():
        missing_fields.append("Teacher/member name")
    if not exam_name.strip():
        missing_fields.append("Examination/program name")
    if not exam_year.strip():
        missing_fields.append("Examination year")

    reference = bill_reference.strip() or (
        f"DUET-EB-{prepared_date.strftime('%Y%m%d')}-{re.sub(r'[^A-Za-z0-9]', '', teacher_name)[:10].upper() or 'TEACHER'}"
    )
    bill_data = {
        "reference": reference,
        "exam_name": exam_name.strip(),
        "session": academic_session.strip(),
        "exam_year": exam_year.strip(),
        "semester": semester,
        "category": category,
        "bill_department": bill_department,
        "teacher": teacher_name.strip(),
        "designation": designation.strip(),
        "home_department": home_department,
        "employee_id": employee_id.strip(),
        "committee_role": committee_role.strip(),
        "teacher_email": teacher_email.strip(),
        "prepared_date": prepared_date.strftime("%d %B %Y"),
        "class_group": class_group.strip(),
        "status": "DRAFT",
    }

    if missing_fields:
        st.warning("Complete these fields before downloading: " + ", ".join(missing_fields))
    else:
        xlsx_bytes = make_excel(bill_data, st.session_state.bill_lines)
        dl1, dl2 = st.columns(2)
        with dl1:
            st.download_button(
                "⬇ Download individual bill (Excel)",
                data=xlsx_bytes,
                file_name=f"{re.sub(r'[^A-Za-z0-9_-]+', '_', teacher_name.strip())}_Exam_Bill.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
        with dl2:
            try:
                pdf_bytes = make_pdf(bill_data, st.session_state.bill_lines)
                st.download_button(
                    "⬇ Download individual bill (PDF)",
                    data=pdf_bytes,
                    file_name=f"{re.sub(r'[^A-Za-z0-9_-]+', '_', teacher_name.strip())}_Exam_Bill.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
            except ImportError:
                st.info("PDF export requires reportlab. Install it using: pip install reportlab")

    if st.button("Clear current individual bill", type="secondary"):
        st.session_state.bill_lines = []
        st.rerun()
else:
    st.info("No bill items have been added. Select an examination duty above and click 'Add this item'.")

st.divider()
st.caption(
    "Developed by Md. Zakirul Hasan | Lecturer, Department of Food Engineering, "
    "DUET, Gazipur. This is an independent utility and does not imply official DUET endorsement."
)
