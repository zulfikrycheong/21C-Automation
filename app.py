import base64
from datetime import datetime
import io
import json
import os
import re

import docx
from google.oauth2.service_account import Credentials
import gspread
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
import streamlit as st

# Force clean full-width canvas layout workspace view
st.set_page_config(page_title="21 Chambers Client List", layout="wide")

if os.path.exists("Company Logo.png"):
    st.image("Company Logo.png", width=250)

st.title("📂 21 Chambers Automator")
st.markdown("##### *Dual-Stream Intake Pipeline & Native Vector PDF Center*")
st.write(
    "Drag and drop your document below to preview, edit if necessary, synchronize the master matrix, and download print-ready layouts."
)
st.markdown("---")

# --- 1. GOOGLE SHEETS MATRIX SYNC ---
GOOGLE_SHEET_NAME = "21 Chambers LLC Client List"
SHEET_TAB_NAME = datetime.now().strftime("%B %Y")
CURRENT_YEAR = datetime.now().strftime("%Y")


def get_google_sheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    if os.path.exists("credentials.json"):
        creds = Credentials.from_service_account_file(
            "credentials.json", scopes=scopes
        )
    else:
        encoded_str = st.secrets["encoded_creds"]
        decoded_bytes = base64.b64decode(encoded_str)
        creds_dict = json.loads(decoded_bytes)
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)

    client = gspread.authorize(creds)
    workbook = client.open(GOOGLE_SHEET_NAME)

    try:
        sheet = workbook.worksheet(SHEET_TAB_NAME)
    except gspread.exceptions.WorksheetNotFound:
        template_sheet = workbook.worksheet("Template")
        sheet = workbook.duplicate_sheet(
            source_sheet_id=template_sheet.id,
            new_sheet_name=SHEET_TAB_NAME,
            insert_sheet_index=0,
        )
        headers = [[
            "Index",
            "Date Opened",
            "File / Matter No.",
            "Matter Type",
            "Client Name(s)",
            "Contact Details",
            "Referral Source",
            "Logged to Matrix",
            "Remarks",
        ]]
        sheet.update(range_name="A1:I1", values=headers)

    return sheet


def get_next_matter_number(sheet):
    try:
        vals = sheet.col_values(3)[1:]
        valid = [int(v.strip()) for v in vals if str(v).strip().isdigit()]
    except Exception:
        valid = []

    matter_nos = sheet.col_values(3)
    mx, mx_idx = -1, 1
    for idx, val in enumerate(matter_nos):
        if str(val).strip().isdigit() and int(str(val).strip()) > mx:
            mx = int(str(val).strip())
            mx_idx = idx + 1

    target_row = mx_idx + 1 if mx != -1 else 2
    base_max = mx if mx != -1 else (max(valid) if valid else 20260728)
    next_no = str(base_max + 1)
    next_idx = target_row - 1
    return next_no, target_row, next_idx


# --- 2. VECTOR PDF ENGINE ---
def generate_perfect_pdf(
    matter_no, clients_text, contacts_text, matter_type, date_opened
):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=21.24,
        rightMargin=21.24,
        topMargin=36.0,
        bottomMargin=36.0,
    )

    style_normal_20 = ParagraphStyle(
        "Norm20",
        fontName="Times-Roman",
        fontSize=20,
        leading=26,
        alignment=TA_LEFT,
    )
    style_bold_20 = ParagraphStyle(
        "Bold20",
        fontName="Times-Bold",
        fontSize=20,
        leading=26,
        alignment=TA_LEFT,
    )
    style_firm_title = ParagraphStyle(
        "FirmTitle",
        fontName="Times-Bold",
        fontSize=20,
        leading=26,
        alignment=TA_CENTER,
    )
    style_firm_body = ParagraphStyle(
        "FirmBody",
        fontName="Times-Roman",
        fontSize=20,
        leading=26,
        alignment=TA_CENTER,
    )
    style_giant_foot = ParagraphStyle(
        "GiantFoot",
        fontName="Helvetica-Bold",
        fontSize=109,
        leading=115,
        alignment=TA_CENTER,
    )

    story = []

    # TOP SECTION
    client_lines = [
        line.strip() for line in clients_text.split("\n") if line.strip()
    ]
    contact_lines = [
        line.strip() for line in contacts_text.split("\n") if line.strip()
    ]

    party_elements = []
    if len(client_lines) > 0:
        party_elements.append(Paragraph(client_lines[0], style_normal_20))
    if len(contact_lines) > 0:
        party_elements.append(Paragraph(contact_lines[0], style_normal_20))

    party_elements.append(Spacer(1, 14))

    if len(client_lines) > 1:
        party_elements.append(Paragraph(client_lines[1], style_normal_20))
    if len(contact_lines) > 1:
        party_elements.append(Paragraph(contact_lines[1], style_normal_20))

    col_w_left = 1.333 * 72
    col_w_right = 6.346 * 72

    top_table = Table(
        [["", party_elements]], colWidths=[col_w_left, col_w_right]
    )
    top_table.setStyle(
        TableStyle([
            ("GRID", (0, 0), (-1, -1), 1.5, colors.black),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("LEFTPADDING", (0, 0), (-1, -1), 7.2),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ])
    )
    story.append(top_table)
    story.append(Spacer(1, 24))

    # CENTRAL HEADER BLOCK
    story.append(Paragraph("21 CHAMBERS LLC", style_firm_title))
    story.append(Spacer(1, 4))
    story.append(
        Paragraph(
            "2 HAVELOCK ROAD #06-17<br/>HAVELOCK 2<br/>SINGAPORE 059763",
            style_firm_body,
        )
    )
    story.append(Spacer(1, 4))
    story.append(
        Paragraph(
            "TEL: 6224 1848 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; FAX: 6223 3092",
            style_firm_body,
        )
    )
    story.append(Spacer(1, 40))

    # BOTTOM MATRIX
    full_matter_name = "Uncontested Divorce"
    if matter_type == "CD":
        full_matter_name = "Contested Divorce"
    elif matter_type == "Annulment":
        full_matter_name = "Annulment"
    elif matter_type == "Variation":
        full_matter_name = "Variation"
    elif matter_type == "Others":
        full_matter_name = "Others"

    file_block_text = (
        f"{matter_no}<br/>Opening date: {date_opened}<br/>Closure date:"
    )

    matrix_rows = [
        [
            Paragraph("SUBJECT<br/>MATTER", style_bold_20),
            Paragraph(full_matter_name, style_normal_20),
        ],
        [
            Paragraph("FILE", style_normal_20),
            Paragraph(file_block_text, style_normal_20),
        ],
        [
            Paragraph("Legal Fee", style_normal_20),
            Paragraph("CASH", style_normal_20),
        ],
        [Paragraph("Remarks", style_normal_20), Paragraph("", style_normal_20)],
    ]

    b_col1 = 1.596 * 72
    b_col2 = 6.083 * 72

    bottom_table = Table(matrix_rows, colWidths=[b_col1, b_col2])
    bottom_table.setStyle(
        TableStyle([
            ("GRID", (0, 0), (-1, -1), 1.5, colors.black),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("LEFTPADDING", (0, 0), (-1, -1), 7.2),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ])
    )
    story.append(bottom_table)
    story.append(Spacer(1, 54))

    # BOTTOM SECTION: Giant Footer
    story.append(Paragraph(matter_no, style_giant_foot))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


# --- 3. DOCX PARSING INTAKE ENGINE ---
def extract_matter_data(doc_path):
    doc = docx.Document(doc_path)
    text_lines = []
    for para in doc.paragraphs:
        if para.text.strip():
            text_lines.append(para.text)
    for table in doc.tables:
        for row in table.rows:
            row_text = [
                cell.text.strip() for cell in row.cells if cell.text.strip()
            ]
            if row_text:
                text_lines.append(" ".join(row_text))

    full_text = "\n".join(text_lines)

    if "uncontested divorce" in full_text.lower():
        matter_type = "UD"
    elif "contested divorce" in full_text.lower():
        matter_type = "CD"
    elif "annulment" in full_text.lower():
        matter_type = "Annulment"
    elif "variation" in full_text.lower():
        matter_type = "Variation"
    else:
        matter_type = "Others"

    app_match = re.search(
        r"Applicant\s*–\s*([^\n\d]+)", full_text, re.IGNORECASE
    )
    res_match = re.search(
        r"Respondent\s*–\s*([^\n\d]+)", full_text, re.IGNORECASE
    )

    applicant_name = (
        re.split(
            r"\s*[-\s–]\s*upload", app_match.group(1), flags=re.IGNORECASE
        )[0]
        .strip()
        .upper()
        if app_match
        else "NIL"
    )
    respondent_name = (
        re.split(
            r"\s*[-\s–]\s*upload", res_match.group(1), flags=re.IGNORECASE
        )[0]
        .strip()
        .upper()
        if res_match
        else "NIL"
    )

    clean_text = full_text.replace("6224 1848", "").replace("6223 3092", "")
    raw_mobiles = re.findall(
        r"\+?\b(?:65|60|1|44)?[ \-]?[89]\d{3}[ \-]?\d{4}\b|\+?60[ \-]?1\d[ \-]?\d{2,3}[ \-]?\d{4}\b|\b01\d[ \-]?\d{2,3}[ \-]?\d{4}\b",
        clean_text,
    )
    mobiles = [re.sub(r"[\s\-+]", "", mob) for mob in raw_mobiles]
    emails = re.findall(
        r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", clean_text
    )

    def format_phone(clean_mob):
        if clean_mob.startswith("01"):
            clean_mob = f"60{clean_mob[1:]}"
        if clean_mob.startswith(("8", "9")) and len(clean_mob) == 8:
            return f"+65 {clean_mob}"
        else:
            return f"+{clean_mob[:2]} {clean_mob[2:]}"

    app_mob = format_phone(mobiles[0]) if len(mobiles) > 0 else "NIL"
    app_email = emails[0] if len(emails) > 0 else "NIL"

    res_mob = format_phone(mobiles[1]) if len(mobiles) > 1 else "NIL"
    res_email = emails[1] if len(emails) > 1 else "NIL"

    referral = "Google Ad"
    if "referral" in full_text.lower():
        ref_match = re.search(
            r"referral\s*[\s,:\-–]\s*(\w+)", full_text, re.IGNORECASE
        )
        if ref_match and "jav" in ref_match.group(1).lower():
            referral = "Javern"
    elif "jav" in full_text.lower():
        referral = "Javern"

    parsed_dict = {
        "matter_type": matter_type,
        "applicant_name": applicant_name,
        "respondent_name": respondent_name,
        "app_mob": app_mob,
        "app_email": app_email,
        "res_mob": res_mob,
        "res_email": res_email,
        "referral": referral,
    }
    return parsed_dict


import pypdfium2 as pdfium


def render_pdf_preview(pdf_bytes):
    """Renders PDF bytes directly as crisp images to avoid browser iframe security blocks."""
    pdf = pdfium.PdfDocument(pdf_bytes)
    for i, page in enumerate(pdf):
        # Render page at 2x scale for sharp text display
        image = page.render(scale=2).to_pil()
        st.image(
            image,
            caption=f"Page {i + 1}",
            use_container_width=True,
        )


# --- 4. STREAMLIT FRAMEWORK FLOWS ---
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0
if "processed_data" not in st.session_state:
    st.session_state["processed_data"] = {}
if "pdf_binary_store" not in st.session_state:
    st.session_state["pdf_binary_store"] = {}
if "synced_records" not in st.session_state:
    st.session_state["synced_records"] = set()

# --- SIDEBAR ACCESS CONFIGURATION PANEL ---
with st.sidebar:
    st.markdown("### 🔑 SYSTEM CONFIGURATION")
    with st.expander("🌐 NEW YEAR / NEW SHEET PROTOCOL", expanded=False):
        st.markdown("""
            ### Connecting a New Master Sheet
            If you are starting a new tracking format or switching to a new Google Sheet file for a new calendar year, the automation machine needs authorization.
            
            **Follow these steps to authorize the link instantly:**
            1. Open your new master Google Sheet.
            2. Click the blue **Share** button (top right).
            3. Invite the system integration email below as an **Editor**:
            """)

        try:
            if os.path.exists("credentials.json"):
                with open("credentials.json") as f:
                    c_data = json.load(f)
                    svc_email = c_data.get(
                        "client_email",
                        "your-service-account@iam.gserviceaccount.com",
                    )
            else:
                encoded_str = st.secrets["encoded_creds"]
                decoded_bytes = base64.b64decode(encoded_str)
                creds_dict = json.loads(decoded_bytes)
                svc_email = creds_dict.get(
                    "client_email",
                    "your-service-account@iam.gserviceaccount.com",
                )
        except Exception:
            svc_email = "your-service-account@iam.gserviceaccount.com"

        st.code(svc_email, language="text")
        st.markdown("""
            4. Uncheck 'Notify people' and click **Share**.
            
            *⚠️ **Note:** If the app throws an 'API Error' during file uploads, this account authorization step was skipped.*
            """)

# --- DYNAMIC THEME CATCH ZONE ---
st.markdown(
    """
    <style>
        [data-testid="stFileUploaderDropzone"] {
            padding: 4rem 2rem !important;
            border-radius: 12px !important;
            background-color: var(--background-color) !important;
            border: 2px dashed var(--text-color) !important;
            opacity: 0.85;
        }
        [data-testid="stFileUploaderDropzone"] label, 
        [data-testid="stFileUploaderDropzone"] div {
            color: var(--text-color) !important;
        }
        [data-testid="stFileUploaderDropzone"] button {
            background-color: var(--secondary-background-color) !important;
            color: var(--text-color) !important;
            border: 1px solid var(--text-color) !important;
        }
    </style>
""",
    unsafe_allow_html=True,
)

uploaded_file = st.file_uploader(
    "Drag and drop Open File Sheet (.docx) here",
    type=["docx"],
    key=f"uploader_{st.session_state['uploader_key']}",
)

if uploaded_file:
    file_id = uploaded_file.name

    # Initialize file processing and load draft sequence numbers
    if file_id not in st.session_state["processed_data"]:
        with st.spinner("Parsing intake parameters and generating preview..."):
            extracted = extract_matter_data(uploaded_file)
            sheet = get_google_sheet()
            next_no, target_row, next_idx = get_next_matter_number(sheet)
            today_str = datetime.now().strftime("%d %B %Y").lstrip("0").upper()

            clients_str = f"APPLICANT - {extracted['applicant_name']}\nRESPONDENT - {extracted['respondent_name']}"
            contacts_str = f"{extracted['app_mob']} {extracted['app_email']}\n{extracted['res_mob']} {extracted['res_email']}".strip()

            initial_pdf = generate_perfect_pdf(
                next_no,
                clients_str,
                contacts_str,
                extracted["matter_type"],
                today_str,
            )

            st.session_state["processed_data"][file_id] = {
                "matter_no": next_no,
                "target_row": target_row,
                "next_idx": next_idx,
                "today_str": today_str,
                "matter_type": extracted["matter_type"],
                "applicant_name": extracted["applicant_name"],
                "respondent_name": extracted["respondent_name"],
                "app_mob": extracted["app_mob"],
                "app_email": extracted["app_email"],
                "res_mob": extracted["res_mob"],
                "res_email": extracted["res_email"],
                "referral": extracted["referral"],
            }
            st.session_state["pdf_binary_store"][file_id] = initial_pdf

    active_data = st.session_state["processed_data"][file_id]

    st.markdown("---")
    st.subheader("📄 Document Preview & Layout Verification")

    # Render Document Preview
    render_pdf_preview(st.session_state["pdf_binary_store"][file_id])

    st.write("")
    c_approve, c_cancel = st.columns([1, 1])

    # Fast Path: Direct Approval
    with c_approve:
        is_synced = file_id in st.session_state["synced_records"]
        if st.button(
            "✅ Approve & Sync to Google Sheets",
            use_container_width=True,
            type="primary",
            disabled=is_synced,
        ):
            with st.spinner("Logging to Google Sheet Master Matrix..."):
                sheet = get_google_sheet()
                cls = f"APPLICANT - {active_data['applicant_name']}\nRESPONDENT - {active_data['respondent_name']}"
                cnt = f"{active_data['app_mob']} {active_data['app_email']}\n{active_data['res_mob']} {active_data['res_email']}".strip()

                new_row = [
                    active_data["next_idx"],
                    active_data["today_str"],
                    active_data["matter_no"],
                    active_data["matter_type"],
                    cls,
                    cnt,
                    active_data["referral"],
                    "Yes",
                    "",
                ]
                t_row = active_data["target_row"]
                sheet.update(range_name=f"A{t_row}:I{t_row}", values=[new_row])

                st.session_state["synced_records"].add(file_id)
                st.toast(
                    f"Synchronized Matrix: Matter {active_data['matter_no']}",
                    icon="🔹",
                )
                st.rerun()

    with c_cancel:
        if st.button("❌ Cancel / Clear", use_container_width=True):
            st.session_state["uploader_key"] += 1
            st.session_state["processed_data"] = {}
            st.session_state["pdf_binary_store"] = {}
            st.session_state["synced_records"] = set()
            st.rerun()

    # Expandable Edit Form
    with st.expander("✏️ Edit Document Details", expanded=False):
        st.caption(
            "Modify any field below to update both the vector PDF cover sheet and the Google Sheet record."
        )

        with st.form("edit_intake_data_form"):
            col1, col2 = st.columns(2)

            with col1:
                edit_matter_no = st.text_input(
                    "Matter File No.", value=active_data["matter_no"]
                )
                edit_date = st.text_input(
                    "Date Opened", value=active_data["today_str"]
                )
                edit_matter_type = st.selectbox(
                    "Matter Type",
                    options=["UD", "CD", "Annulment", "Variation", "Others"],
                    index=["UD", "CD", "Annulment", "Variation", "Others"].index(
                        active_data["matter_type"]
                    )
                    if active_data["matter_type"]
                    in ["UD", "CD", "Annulment", "Variation", "Others"]
                    else 0,
                )
                edit_referral = st.text_input(
                    "Referral Source", value=active_data["referral"]
                )

            with col2:
                edit_app_name = st.text_input(
                    "Applicant Name", value=active_data["applicant_name"]
                )
                edit_app_mob = st.text_input(
                    "Applicant Contact Number", value=active_data["app_mob"]
                )
                edit_app_email = st.text_input(
                    "Applicant Email", value=active_data["app_email"]
                )

                edit_res_name = st.text_input(
                    "Respondent Name", value=active_data["respondent_name"]
                )
                edit_res_mob = st.text_input(
                    "Respondent Contact Number", value=active_data["res_mob"]
                )
                edit_res_email = st.text_input(
                    "Respondent Email", value=active_data["res_email"]
                )

            save_submit = st.form_submit_button(
                "💾 Save & Update Cover Sheet", use_container_width=True
            )

        if save_submit:
            # Update state parameters
            active_data.update({
                "matter_no": edit_matter_no,
                "today_str": edit_date,
                "matter_type": edit_matter_type,
                "referral": edit_referral,
                "applicant_name": edit_app_name,
                "app_mob": edit_app_mob,
                "app_email": edit_app_email,
                "respondent_name": edit_res_name,
                "res_mob": edit_res_mob,
                "res_email": edit_res_email,
            })

            # Recompile vector PDF with clean formatting strings
            cls = f"APPLICANT - {edit_app_name}\nRESPONDENT - {edit_res_name}"
            cnt = f"{edit_app_mob} {edit_app_email}\n{edit_res_mob} {edit_res_email}".strip()

            new_pdf_bytes = generate_perfect_pdf(
                edit_matter_no, cls, cnt, edit_matter_type, edit_date
            )
            st.session_state["pdf_binary_store"][file_id] = new_pdf_bytes

            # Auto-sync revised inputs to Google Sheet
            with st.spinner("Syncing updated details to Google Sheets..."):
                sheet = get_google_sheet()
                new_row = [
                    active_data["next_idx"],
                    edit_date,
                    edit_matter_no,
                    edit_matter_type,
                    cls,
                    cnt,
                    edit_referral,
                    "Yes",
                    "",
                ]
                t_row = active_data["target_row"]
                sheet.update(range_name=f"A{t_row}:I{t_row}", values=[new_row])
                st.session_state["synced_records"].add(file_id)

            st.toast("Updated and re-synced successfully!", icon="✅")
            st.rerun()

    # Download Trigger
    if file_id in st.session_state["synced_records"]:
        st.markdown("---")
        m_tag = active_data["matter_no"].replace("/", "_").strip()
        st.download_button(
            label=f"🖨️ Download Final Cover Sheet PDF (Matter File Number: {m_tag})",
            data=st.session_state["pdf_binary_store"][file_id],
            file_name=f"21Chambers_Cover_{m_tag}.pdf",
            mime="application/pdf",
            use_container_width=True,
            type="primary",
        )
