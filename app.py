import base64
from datetime import datetime
import io
import json
import os
import re
import sqlite3

import docx
from google.oauth2.service_account import Credentials
import gspread
import pypdfium2 as pdfium
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
import streamlit as st
import streamlit.components.v1 as components

# --- APP CONFIGURATION & BRANDING ---
st.set_page_config(
    page_title="ChambersOS | Practice Operations Suite",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Custom Executive Theme Injection (Eggshell & Matte Charcoal)
st.markdown(
    """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
        
        /* Global Eggshell Background & Matte Charcoal Text */
        html, body, [class*="css"] {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
            color: #1e2229;
        }
        
        .stApp {
            background-color: #fcfbf9 !important;
        }
        
        .main-header {
            padding: 0.8rem 0rem 1.4rem 0rem;
            border-bottom: 1px solid #e2e8f0;
            margin-bottom: 1.5rem;
        }
        
        .suite-title {
            font-size: 1.85rem;
            font-weight: 800;
            color: #1e2229;
            letter-spacing: -0.5px;
            margin: 0;
        }
        
        .suite-subtitle {
            font-size: 0.95rem;
            color: #64748b;
            font-weight: 500;
            margin-top: 4px;
        }
        
        /* Clean Tab Overhaul — Eliminates Ugly Purple Focus/Click Artifacts */
        .stTabs [data-baseweb="tab-list"] {
            gap: 12px;
            border-bottom: 2px solid #e2e8f0;
            padding-bottom: 2px;
            overflow: visible !important;
            background-color: transparent !important;
        }
        
        .stTabs [data-baseweb="tab"] {
            height: 48px;
            white-space: nowrap !important;
            border-radius: 8px 8px 0px 0px;
            padding: 10px 22px !important;
            font-weight: 600;
            font-size: 0.92rem;
            color: #64748b;
            border: none !important;
            background-color: transparent !important;
            overflow: visible !important;
            outline: none !important;
            box-shadow: none !important;
        }
        
        .stTabs [data-baseweb="tab"]:focus, 
        .stTabs [data-baseweb="tab"]:active {
            color: #1e40af !important;
            background-color: transparent !important;
            outline: none !important;
            box-shadow: none !important;
        }
        
        .stTabs [data-baseweb="tab"] div {
            overflow: visible !important;
            color: inherit !important;
        }
        
        .stTabs [aria-selected="true"] {
            color: #1e40af !important;
            border-bottom: 3px solid #1e40af !important;
            background-color: #f4f2ee !important;
            outline: none !important;
            box-shadow: none !important;
        }
        
        /* Fix Metric Truncation & Matte Styling */
        [data-testid="stMetricValue"] {
            font-size: 1.65rem !important;
            white-space: normal !important;
            word-break: break-word !important;
            color: #1e2229 !important;
        }
        
        /* Upload Area Styling */
        [data-testid="stFileUploaderDropzone"] {
            padding: 3.5rem 2rem !important;
            border-radius: 12px !important;
            background-color: #f4f2ee !important;
            border: 2px dashed #cbd5e1 !important;
            transition: all 0.2s ease-in-out;
        }
        [data-testid="stFileUploaderDropzone"]:hover {
            border-color: #1e40af !important;
            background-color: #eff6ff !important;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# --- WORKSPACE TOP BAR ---
c_logo, c_title = st.columns([1, 4])
with c_logo:
    if os.path.exists("Company Logo.png"):
        st.image("Company Logo.png", width=220)
with c_title:
    st.markdown(
        """
        <div class="main-header">
            <h1 class="suite-title">⚖️ ChambersOS Legal Automation Hub</h1>
            <div class="suite-subtitle">Integrated Matter Intake Engine & Crown Worldwide Archival Terminal</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

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

  # Bumped size from 109 to 125 (~15% increase) for easier cutting
  style_giant_foot = ParagraphStyle(
      "GiantFoot",
      fontName="Helvetica-Bold",
      fontSize=125,
      leading=132,
      alignment=TA_CENTER,
  )

  story = []

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

  # LOCKED LAYOUT BLOCK: Nests bottom table and giant number with exact 2pt (2px visual equivalent) spacing
  locked_block_table = Table(
      [[bottom_table], [Paragraph(matter_no, style_giant_foot)]],
      colWidths=[7.942 * 72],
  )
  locked_block_table.setStyle(
      TableStyle([
          ("VALIGN", (0, 0), (-1, -1), "TOP"),
          ("TOPPADDING", (0, 0), (0, 0), 0),
          ("BOTTOMPADDING", (0, 0), (0, 0), 2),  # Exactly 2pt gap below table
          ("TOPPADDING", (0, 1), (0, 1), 2),
          ("BOTTOMPADDING", (0, 1), (0, 1), 0),
          ("LEFTPADDING", (0, 0), (-1, -1), 0),
          ("RIGHTPADDING", (0, 0), (-1, -1), 0),
      ])
  )

  story.append(locked_block_table)

  doc.build(story)
  buffer.seek(0)
  return buffer.getvalue()


# --- 3. DOCX PARSER ENGINE ---
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

    return {
        "matter_type": matter_type,
        "applicant_name": applicant_name,
        "respondent_name": respondent_name,
        "app_mob": app_mob,
        "app_email": app_email,
        "res_mob": res_mob,
        "res_email": res_email,
        "referral": referral,
    }


def render_pdf_preview(pdf_bytes):
    pdf = pdfium.PdfDocument(pdf_bytes)
    for i, page in enumerate(pdf):
        image = page.render(scale=2).to_pil()
        st.image(image, caption=f"Page {i + 1}", use_container_width=True)


# --- INITIALIZE SESSION STATE ---
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0
if "processed_data" not in st.session_state:
    st.session_state["processed_data"] = {}
if "pdf_binary_store" not in st.session_state:
    st.session_state["pdf_binary_store"] = {}
if "synced_records" not in st.session_state:
    st.session_state["synced_records"] = set()

# --- SIDEBAR CONFIGURATION ---
with st.sidebar:
    st.markdown("### ⚙️ SYSTEM SETTINGS")
    with st.expander("🌐 MASTER SHEET SETUP", expanded=False):
        st.markdown("""
            **Connecting a New Master Sheet:**
            1. Open your target Google Sheet.
            2. Share the file with **Editor** permissions to:
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
        st.caption("Ensure 'Notify people' is unchecked before sharing.")

# --- DUAL-WING TO TRIPLE-WING SUITE TABS ---
tab_intake, tab_crown, tab_locator = st.tabs([
    "📥 Intake Pipeline & Cover Generator",
    "📦 Crown Box Archival Terminal",
    "🔍 Warehouse Archive Locator",
])

# ==============================================================================
# WING 1: INTAKE PIPELINE & COVER GENERATOR
# ==============================================================================
with tab_intake:
    st.markdown(
        "Drop in your completed **Open File Sheet (`.docx`)** to auto-index against the master matrix and generate the vector cover sheet."
    )

    uploaded_file = st.file_uploader(
        "Drop Open File Sheet (.docx) here",
        type=["docx"],
        key=f"uploader_{st.session_state['uploader_key']}",
    )

    if uploaded_file:
        file_id = uploaded_file.name

        if file_id not in st.session_state["processed_data"]:
            with st.spinner(
                "Parsing intake parameters and synchronizing sequence..."
            ):
                extracted = extract_matter_data(uploaded_file)
                sheet = get_google_sheet()
                next_no, target_row, next_idx = get_next_matter_number(sheet)
                today_str = (
                    datetime.now().strftime("%d %B %Y").lstrip("0").upper()
                )

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
        st.subheader("📄 Generated Vector Cover Preview")
        render_pdf_preview(st.session_state["pdf_binary_store"][file_id])

        st.write("")
        c_approve, c_cancel = st.columns([1, 1])

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
                    sheet.update(
                        range_name=f"A{t_row}:I{t_row}", values=[new_row]
                    )

                    st.session_state["synced_records"].add(file_id)
                    st.toast(
                        f"Matrix Synchronized: File {active_data['matter_no']}",
                        icon="🔹",
                    )
                    st.rerun()

        with c_cancel:
            if st.button("❌ Clear Session", use_container_width=True):
                st.session_state["uploader_key"] += 1
                st.session_state["processed_data"] = {}
                st.session_state["pdf_binary_store"] = {}
                st.session_state["synced_records"] = set()
                st.rerun()

        with st.expander("✏️ Modify Matter Parameters", expanded=False):
            with st.form("edit_intake_data_form"):
                col1, col2 = st.columns(2)

                with col1:
                    edit_matter_no = st.text_input(
                        "File / Matter No.", value=active_data["matter_no"]
                    )
                    edit_date = st.text_input(
                        "Opening Date", value=active_data["today_str"]
                    )
                    edit_matter_type = st.selectbox(
                        "Matter Classification",
                        options=[
                            "UD",
                            "CD",
                            "Annulment",
                            "Variation",
                            "Others",
                        ],
                        index=[
                            "UD",
                            "CD",
                            "Annulment",
                            "Variation",
                            "Others",
                        ].index(active_data["matter_type"])
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
                        "Applicant Contact Number",
                        value=active_data["app_mob"],
                    )
                    edit_app_email = st.text_input(
                        "Applicant Email", value=active_data["app_email"]
                    )
                    edit_res_name = st.text_input(
                        "Respondent Name", value=active_data["respondent_name"]
                    )
                    edit_res_mob = st.text_input(
                        "Respondent Contact Number",
                        value=active_data["res_mob"],
                    )
                    edit_res_email = st.text_input(
                        "Respondent Email", value=active_data["res_email"]
                    )

                save_submit = st.form_submit_button(
                    "💾 Save & Update Parameters", use_container_width=True
                )

            if save_submit:
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

                cls = f"APPLICANT - {edit_app_name}\nRESPONDENT - {edit_res_name}"
                cnt = f"{edit_app_mob} {edit_app_email}\n{edit_res_mob} {edit_res_email}".strip()

                new_pdf_bytes = generate_perfect_pdf(
                    edit_matter_no, cls, cnt, edit_matter_type, edit_date
                )
                st.session_state["pdf_binary_store"][file_id] = new_pdf_bytes

                with st.spinner("Synchronizing revisions to Google Sheets..."):
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
                    sheet.update(
                        range_name=f"A{t_row}:I{t_row}", values=[new_row]
                    )
                    st.session_state["synced_records"].add(file_id)

                st.toast("Cover sheet regenerated & synced!", icon="✅")
                st.rerun()

        if file_id in st.session_state["synced_records"]:
            st.markdown("---")
            m_tag = active_data["matter_no"].replace("/", "_").strip()
            st.download_button(
                label=f"🖨️ Download Print-Ready Cover Sheet (Matter: {m_tag})",
                data=st.session_state["pdf_binary_store"][file_id],
                file_name=f"21Chambers_Cover_{m_tag}.pdf",
                mime="application/pdf",
                use_container_width=True,
                type="primary",
            )

# ==============================================================================
# WING 2: CROWN BOX ARCHIVAL TERMINAL (WITH LOCALSTORAGE QUEUE PERSISTENCE)
# ==============================================================================
with tab_crown:
    crown_scanner_component = """
    <!DOCTYPE html>
    <html lang="en">
    <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <script src="https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js"></script>
    <script src="https://unpkg.com/docx@8.5.0/build/index.umd.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js"></script>
    <style>
      :root {
        --primary: #1e40af;
        --primary-hover: #1d4ed8;
        --success: #059669;
        --warning: #d97706;
        --danger: #dc2626;
        --surface: #ffffff;
        --background: #f8fafc;
        --border: #e2e8f0;
        --text: #0f172a;
        --text-muted: #64748b;
      }
      * { box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }
      body { background: var(--background); color: var(--text); padding: 10px; }

      .card {
        background: var(--surface);
        border-radius: 10px;
        border: 1px solid var(--border);
        padding: 16px;
        margin-bottom: 14px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
      }
      .form-label { font-size: 0.82rem; font-weight: 700; color: var(--text-muted); margin-bottom: 5px; display: block; text-transform: uppercase; letter-spacing: 0.5px; }
      .form-control {
        width: 100%;
        padding: 10px 12px;
        font-size: 0.95rem;
        border: 1px solid #cbd5e1;
        border-radius: 6px;
        background: #fff;
        margin-bottom: 12px;
        color: var(--text);
      }
      .form-control:focus { outline: none; border-color: var(--primary); }

      .capacity-container { margin-bottom: 12px; }
      .capacity-text { display: flex; justify-content: space-between; font-size: 0.8rem; font-weight: 700; margin-bottom: 4px; color: var(--text-muted); }
      .progress-bg { width: 100%; height: 8px; background: #e2e8f0; border-radius: 4px; overflow: hidden; }
      .progress-fill { height: 100%; width: 0%; background: var(--primary); transition: width 0.2s ease; }

      .camera-box { display: none; background: #000; border-radius: 10px; overflow: hidden; margin-bottom: 14px; }
      .video-wrapper { position: relative; width: 100%; background: #000; }
      video { width: 100%; height: auto; display: block; }
      .camera-frame { position: absolute; top: 8%; left: 6%; right: 6%; bottom: 8%; pointer-events: none; border: 2px dashed rgba(255, 255, 255, 0.5); border-radius: 6px; }
      .camera-controls { padding: 12px; background: #0f172a; display: flex; gap: 8px; }

      .btn-row { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 12px; }
      .btn { padding: 12px 14px; font-size: 0.92rem; font-weight: 600; border: none; border-radius: 6px; cursor: pointer; text-align: center; display: inline-flex; align-items: center; justify-content: center; gap: 8px; }
      .btn-primary { background: var(--primary); color: #fff; }
      .btn-success { background: var(--success); color: #fff; }
      .btn-secondary { background: #f1f5f9; color: var(--text); border: 1px solid #cbd5e1; }
      .btn-block { width: 100%; }

      .badge-bar { display: flex; gap: 6px; flex-wrap: wrap; margin-bottom: 8px; }
      .badge-btn { background: #e2e8f0; color: #1e293b; border: none; padding: 4px 8px; border-radius: 4px; font-size: 0.78rem; font-weight: 600; cursor: pointer; }
      .modifier-btn { background: #f1f5f9; border: 1px solid #cbd5e1; color: #475569; padding: 2px 6px; border-radius: 4px; font-size: 0.75rem; font-weight: 600; cursor: pointer; }

      .queue-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; }
      .queue-count { font-size: 0.88rem; font-weight: 700; color: var(--text-muted); }
      .file-item {
        background: #fff; border: 1px solid var(--border); border-left: 4px solid var(--primary); border-radius: 6px; padding: 10px 12px; margin-bottom: 8px; font-size: 0.88rem; position: relative; cursor: pointer; line-height: 1.45;
      }
      .file-item .line-label { font-weight: 700; color: #475569; }
      .file-item .file-no { font-weight: 700; color: var(--primary); }
      .file-item .matter-type { font-weight: 600; color: var(--success); }
      .file-item .del-btn { position: absolute; right: 8px; top: 8px; background: none; border: none; color: var(--danger); cursor: pointer; font-size: 1.2rem; padding: 2px 8px; }

      .modal-backdrop { position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(15, 23, 42, 0.6); display: none; justify-content: center; align-items: flex-end; z-index: 100; }
      .modal-card { background: #fff; width: 100%; max-width: 580px; border-radius: 16px 16px 0 0; padding: 20px 16px; max-height: 90vh; overflow-y: auto; }
      .ocr-banner { background: #eff6ff; color: var(--primary); padding: 10px 14px; border-radius: 6px; font-size: 0.85rem; font-weight: 500; margin-bottom: 14px; display: flex; align-items: center; gap: 10px; }
      .spinner { border: 2px solid #bfdbfe; border-top: 2px solid var(--primary); border-radius: 50%; width: 14px; height: 14px; animation: spin 0.7s linear infinite; }
      @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
    </style>
    </head>
    <body>

    <div class="card">
      <label class="form-label">Carton Number (Box Identifier)</label>
      <input type="text" id="cartonNumberInput" class="form-control" placeholder="e.g. YYLEE – 21C - 282- YX" oninput="saveCartonState()">
      
      <label class="form-label">Active Archive Department</label>
      <select id="departmentSelect" class="form-control" onchange="changeDepartment()">
        <option value="CONVEYANCING" selected>Conveyancing & Real Estate (YY Lee / JL / SH)</option>
        <option value="FAMILY">Family & Matrimonial (21 Chambers / YX / YT)</option>
        <option value="PROBATE">Probate & Estate (SL / KYE)</option>
      </select>

      <div class="capacity-container">
        <div class="capacity-text">
          <span id="capacityStatusText">Carton Fill Capacity</span>
          <span id="capacityRatio">0 / 30</span>
        </div>
        <div class="progress-bg"><div class="progress-fill" id="progressBar"></div></div>
      </div>
    </div>

    <div class="card">
      <div class="btn-row">
        <button class="btn btn-primary" onclick="openCamera()">Live Camera Scan</button>
        <button class="btn btn-secondary" onclick="openManualEntry()">Manual Entry</button>
      </div>
      <input type="file" id="filePicker" accept="image/*" style="display:none" onchange="handleImageFile(this.files[0])">
      <button class="btn btn-secondary btn-block" onclick="document.getElementById('filePicker').click()">Upload File Cover Photo</button>

      <div class="camera-box" id="cameraBox">
        <div class="video-wrapper">
          <video id="videoElement" autoplay playsinline></video>
          <div class="camera-frame"></div>
        </div>
        <div class="camera-controls">
          <button class="btn btn-success" style="flex:1" onclick="snapPhoto()">Snap & Parse</button>
          <button class="btn btn-secondary" style="background:#334155; color:#fff; border:none;" onclick="closeCamera()">Cancel</button>
        </div>
      </div>
      <canvas id="snapshotCanvas" style="display:none"></canvas>
    </div>

    <div class="card">
      <div class="queue-header">
        <div class="queue-count" id="queueCount">0 Files in Active Box</div>
        <button class="btn btn-success" style="padding: 8px 14px; font-size: 0.85rem;" onclick="generateWordDoc()">Export .docx Record</button>
      </div>
      <div id="fileQueueContainer"></div>
    </div>

    <!-- Edit Modal -->
    <div class="modal-backdrop" id="editModal">
      <div class="modal-card">
        <div id="ocrStatus" class="ocr-banner" style="display:none"><div class="spinner"></div> <span>Parsing physical jacket...</span></div>
        <h2 style="font-size: 1.15rem; margin-bottom: 14px;">Review & Standardize Record</h2>

        <label class="form-label">
          1. File Number
          <span style="float: right; text-transform: none; font-weight: normal;">
            Tags: 
            <button type="button" class="modifier-btn" onclick="toggleTag('[ABORTED]')">[ABORTED]</button>
            <button type="button" class="modifier-btn" onclick="toggleTag('(Vol 1)')">Vol 1</button>
            <button type="button" class="modifier-btn" onclick="toggleTag('(Vol 2)')">Vol 2</button>
          </span>
        </label>
        <input type="text" id="editLine1" class="form-control" placeholder="e.g. 20260235" onkeydown="handleModalKeyDown(event)">

        <label class="form-label">2. Client / Represented Party</label>
        <input type="text" id="editLine2" class="form-control" placeholder="Client name (Role in parentheticals)" onkeydown="handleModalKeyDown(event)">

        <label class="form-label">3. Master Taxonomy</label>
        <select id="editLine3Select" class="form-control" onchange="handleMatterChange(this.value)" onkeydown="handleModalKeyDown(event)"></select>
        
        <div id="customMatterDiv" style="display: none; margin-top: -6px; margin-bottom: 12px;">
          <input type="text" id="editLine3Custom" class="form-control" placeholder="Enter custom matter description..." onkeydown="handleModalKeyDown(event)">
        </div>

        <label class="form-label" id="editLine4Label">4. Target Metadata</label>
        <div class="badge-bar" id="caseShortcuts" style="display:none;">
          <button type="button" class="badge-btn" onclick="quickFillPrefix('FC/OAD ')">FC/OAD</button>
          <button type="button" class="badge-btn" onclick="quickFillPrefix('FC/OADV ')">FC/OADV</button>
          <button type="button" class="badge-btn" onclick="quickFillPrefix('FC/D ')">FC/D</button>
          <button type="button" class="badge-btn" onclick="quickFillPrefix('FC/P ')">FC/P</button>
          <button type="button" class="badge-btn" onclick="quickFillPrefix('-')">None (-)</button>
        </div>
        <div class="badge-bar" id="conveyancingShortcuts" style="display:flex;">
          <button type="button" class="badge-btn" onclick="lookupPostalCode()">Lookup Postal Code</button>
          <button type="button" class="badge-btn" onclick="quickFillPrefix('Lot No. ')">Lot No.</button>
          <button type="button" class="badge-btn" onclick="document.getElementById('editLine4').value='-'">None (-)</button>
        </div>
        <textarea id="editLine4" class="form-control" rows="2" placeholder="Property Address or Case Number" onkeydown="handleModalKeyDown(event)"></textarea>

        <div class="btn-row" style="margin-top: 16px;">
          <button class="btn btn-secondary" onclick="closeModal()">Discard</button>
          <button class="btn btn-success" id="doneAppendBtn" onclick="saveRecordFromModal()">Done & Append ↵</button>
        </div>
      </div>
    </div>

    <script>
      const TAXONOMY_OPTIONS = {
        CONVEYANCING: [
          "Purchase", "Sale", "BUC Purchase", "HDB Purchase", "Purchase (JTC)",
          "Sub-Purchase", "Decoupling", "Transfer Pursuant to Divorce",
          "Transfer by Way of Assent", "Transmission (TAD)", "Notice of Death (NOD)",
          "Unilateral Severance of Joint Tenancy", "Mutual Severance of Joint Tenancy",
          "Redemption", "Refinancing", "Mortgage", "Discharge of CPF Charge",
          "Replacement of Title (CT/SSCT)", "Lodgement of Caveat",
          "Registration of Lease (JTC)", "Variation of Lease", "OTHER"
        ],
        FAMILY: [
          "UD", "CD", "UA", "CA", "UV", "CV", "Annulment", "Variation",
          "Deed of Separation", "Pre-Nuptial Agreement", "Post-Nuptial Agreement",
          "Settlement Agreement", "PPO", "MSS", "OTHER"
        ],
        PROBATE: [
          "Estate (Probate)", "Estate (GLA)", "Grant of Probate (GOP)", "Letters of Administration (GLA)",
          "Resealing of Grant", "Transmission (TAD)", "Transfer by Way of Assent", "Amendment of Schedule of Assets (SOA)", "OTHER"
        ]
      };

      const CAPACITY_LIMITS = {
        CONVEYANCING: { limit: 30, warn: 25 },
        FAMILY: { limit: 80, warn: 75 },
        PROBATE: { limit: 35, warn: 30 }
      };

      let fileQueue = [];
      let currentEditingIndex = -1;
      let videoStream = null;

      window.addEventListener('DOMContentLoaded', () => {
        loadCartonState();
        changeDepartment();
        updateCapacityDisplay();
      });

      function saveCartonState() {
        const cartonNo = document.getElementById('cartonNumberInput').value;
        localStorage.setItem('chambersos_active_carton', cartonNo);
        localStorage.setItem('chambersos_active_queue', JSON.stringify(fileQueue));
      }

      function loadCartonState() {
        const savedCarton = localStorage.getItem('chambersos_active_carton');
        const savedQueue = localStorage.getItem('chambersos_active_queue');
        if (savedCarton) document.getElementById('cartonNumberInput').value = savedCarton;
        if (savedQueue) {
          try { fileQueue = JSON.parse(savedQueue); renderQueue(); } catch(e) { fileQueue = []; }
        }
      }

      function changeDepartment() {
        const dept = document.getElementById('departmentSelect').value;
        const select = document.getElementById('editLine3Select');
        select.innerHTML = "";

        TAXONOMY_OPTIONS[dept].forEach(opt => {
          const el = document.createElement('option');
          el.value = opt;
          el.textContent = (opt === "OTHER") ? "Other (Specify Custom Matter)..." : opt;
          select.appendChild(el);
        });

        const l4 = document.getElementById('editLine4Label');
        const cShort = document.getElementById('caseShortcuts');
        const pShort = document.getElementById('conveyancingShortcuts');
        if (dept === "CONVEYANCING") {
          l4.textContent = "4. Property Address / Land Lot";
          document.getElementById('editLine4').placeholder = "e.g. 17 Fernvale Close #16-31 Riverbank @ Fernvale, Singapore 797478";
          cShort.style.display = "none";
          pShort.style.display = "flex";
        } else {
          l4.textContent = "4. Case No. (or '-' if unfiled)";
          document.getElementById('editLine4').placeholder = "e.g. FC/OAD 1909/2026 or -";
          cShort.style.display = "flex";
          pShort.style.display = "none";
        }
        updateCapacityDisplay();
      }

      function updateCapacityDisplay() {
        const dept = document.getElementById('departmentSelect').value;
        const { limit, warn } = CAPACITY_LIMITS[dept];
        const count = fileQueue.length;
        const percent = Math.min(100, Math.round((count / limit) * 100));

        document.getElementById('capacityRatio').textContent = `${count} / ${limit}`;
        const bar = document.getElementById('progressBar');
        bar.style.width = percent + "%";

        if (count >= limit) {
          bar.style.backgroundColor = "var(--danger)";
          document.getElementById('capacityStatusText').textContent = "⚠️ CARTON FULL";
          document.getElementById('capacityStatusText').style.color = "var(--danger)";
        } else if (count >= warn) {
          bar.style.backgroundColor = "var(--warning)";
          document.getElementById('capacityStatusText').textContent = "Carton Approaching Limit";
          document.getElementById('capacityStatusText').style.color = "var(--warning)";
        } else {
          bar.style.backgroundColor = "var(--primary)";
          document.getElementById('capacityStatusText').textContent = "Carton Fill Capacity";
          document.getElementById('capacityStatusText').style.color = "var(--text-muted)";
        }
      }

      function toggleTag(tag) {
        const l1 = document.getElementById('editLine1');
        l1.value = l1.value.includes(tag) ? l1.value.replace(tag, '').trim() : `${l1.value.trim()} ${tag}`.trim();
      }

      function quickFillPrefix(prefix) {
        const input = document.getElementById('editLine4');
        if (prefix === "-") { input.value = "-"; return; }
        if (document.getElementById('departmentSelect').value === "CONVEYANCING") {
          input.value = prefix;
          input.focus();
        } else {
          input.value = `${prefix}/${new Date().getFullYear()}`;
          input.focus();
          const pos = prefix.length;
          input.setSelectionRange(pos, pos);
        }
      }

      async function lookupPostalCode() {
        const input = document.getElementById('editLine4');
        const match = input.value.trim().match(/\\b\\d{6}\\b/);
        let postal = match ? match[0] : prompt("Enter 6-digit Singapore Postal Code:");
        if (postal && /^\\d{6}$/.test(postal.trim())) {
          try {
            const res = await fetch(`https://www.onemap.gov.sg/api/common/elastic/search?searchVal=${postal}&returnGeom=N&getAddrDetails=Y`);
            const data = await res.json();
            if (data.results && data.results.length > 0) {
              const item = data.results[0];
              const blk = item.BLK_NO ? `Blk ${item.BLK_NO} ` : '';
              const road = item.ROAD_NAME || '';
              const bldg = (item.BUILDING && item.BUILDING !== 'N/A') ? ` ${item.BUILDING}` : '';
              input.value = `${blk}${road}${bldg}, Singapore ${postal}`.trim();
            } else { alert("No OneMap record for postal " + postal); }
          } catch(e) { alert("Postal lookup unavailable."); }
        }
      }

      function handleMatterChange(val) {
        document.getElementById('customMatterDiv').style.display = (val === "OTHER") ? "block" : "none";
        if (val === "OTHER") document.getElementById('editLine3Custom').focus();
      }

      function handleModalKeyDown(event) {
        if (event.key === "Enter" && event.target.tagName !== "TEXTAREA") {
          event.preventDefault();
          saveRecordFromModal();
        }
      }

      async function openCamera() {
        const box = document.getElementById('cameraBox');
        box.style.display = 'block';
        try {
          videoStream = await navigator.mediaDevices.getUserMedia({
            video: { facingMode: "environment", width: { ideal: 1920 }, height: { ideal: 1080 } }, audio: false
          });
          document.getElementById('videoElement').srcObject = videoStream;
        } catch (err) {
          alert("Camera access denied or unavailable: " + err.message);
          box.style.display = 'none';
        }
      }

      function closeCamera() {
        if (videoStream) { videoStream.getTracks().forEach(t => t.stop()); videoStream = null; }
        document.getElementById('cameraBox').style.display = 'none';
      }

      function snapPhoto() {
        const video = document.getElementById('videoElement');
        const canvas = document.getElementById('snapshotCanvas');
        canvas.width = video.videoWidth || 1280;
        canvas.height = video.videoHeight || 720;
        const ctx = canvas.getContext('2d');
        ctx.drawImage(video, 0, 0, canvas.width, canvas.height);
        canvas.toBlob(blob => { closeCamera(); processImageWithOCR(blob); }, 'image/jpeg', 0.95);
      }

      function handleImageFile(file) { if (file) processImageWithOCR(file); }

      async function processImageWithOCR(imageBlob) {
        openModal(-1);
        const banner = document.getElementById('ocrStatus');
        banner.style.display = 'flex';
        try {
          const worker = await Tesseract.createWorker('eng');
          const ret = await worker.recognize(imageBlob);
          await worker.terminate();
          parseAnchoredCover(ret.data.text);
        } catch (err) {
          alert("OCR parsing failed: " + err.message);
        } finally {
          banner.style.display = 'none';
          document.getElementById('editLine4').focus();
        }
      }

      function parseAnchoredCover(rawText) {
        const FIRM_BLACKLIST = ["059763", "62241848", "62233092", "6224", "6223", "06-17", "21 CHAMBERS", "YY LEE", "HAVELOCK"];
        const dept = document.getElementById('departmentSelect').value;

        let fileNo = "";
        const matches = rawText.match(/\\b20[12]\\d{5,6}\\b|\\bLP\\.\\d{6,8}\\b/g) || [];
        const valid = matches.filter(n => !FIRM_BLACKLIST.includes(n));
        if (valid.length > 0) {
          fileNo = valid[valid.length - 1];
        } else {
          const fallback = rawText.match(/FILE[^\\d]{0,25}(\\d{7,8})/i);
          if (fallback && !FIRM_BLACKLIST.includes(fallback[1])) fileNo = fallback[1];
        }

        const upper = rawText.toUpperCase();
        let matchedMatter = (dept === "CONVEYANCING") ? "Purchase" : "UD";

        if (dept === "CONVEYANCING") {
          if (upper.includes("DECOUPLING")) matchedMatter = "Decoupling";
          else if (upper.includes("TRANSFER PURSUANT TO DIVORCE") || (upper.includes("TRANSFER") && upper.includes("DIVORCE"))) matchedMatter = "Transfer Pursuant to Divorce";
          else if (upper.includes("ASSENT")) matchedMatter = "Transfer by Way of Assent";
          else if (upper.includes("TRANSMISSION") || upper.includes("TAD")) matchedMatter = "Transmission (TAD)";
          else if (upper.includes("NOTICE OF DEATH") || upper.includes("NOD")) matchedMatter = "Notice of Death (NOD)";
          else if (upper.includes("SEVERANCE")) matchedMatter = upper.includes("MUTUAL") ? "Mutual Severance of Joint Tenancy" : "Unilateral Severance of Joint Tenancy";
          else if (upper.includes("BUC")) matchedMatter = "BUC Purchase";
          else if (upper.includes("HDB PURCHASE")) matchedMatter = "HDB Purchase";
          else if (upper.includes("SALE")) matchedMatter = "Sale";
          else if (upper.includes("REDEMPTION")) matchedMatter = "Redemption";
          else if (upper.includes("REFINANCING")) matchedMatter = "Refinancing";
          else if (upper.includes("MORTGAGE")) matchedMatter = "Mortgage";
          else if (upper.includes("CPF")) matchedMatter = "Discharge of CPF Charge";
          else if (upper.includes("CAVEAT")) matchedMatter = "Lodgement of Caveat";
          else if (upper.includes("PURCHASE")) matchedMatter = "Purchase";
        } else {
          if (upper.includes("CONTESTED DIVORCE")) matchedMatter = "CD";
          else if (upper.includes("UNCONTESTED DIVORCE") || upper.includes("UNCONTESTED")) matchedMatter = "UD";
          else if (upper.includes("ANNULMENT")) matchedMatter = "UA";
          else if (upper.includes("VARIATION")) matchedMatter = "UV";
        }

        let client = "";
        const lines = rawText.split('\\n').map(l => l.trim()).filter(l => l.length > 0);
        for (let line of lines) {
          const partyMatch = line.match(/(?:APPLICANT|PURCHASER|BUYER|VENDOR|SELLER|TRANSFEREE|TRANSFEROR|CLIENT)\\s*[-–:]?\\s*(.+)/i);
          if (partyMatch && partyMatch[1]) {
            let roleMatch = line.match(/(PURCHASER|BUYER|VENDOR|SELLER|TRANSFEREE|TRANSFEROR|APPLICANT)/i);
            let cleanedName = cleanString(partyMatch[1]).replace(/^[\\[\\]{}|#\\d\\s]+/, '').trim();
            if (cleanedName.length > 2) {
              client = roleMatch ? `${cleanedName} (${roleMatch[0].charAt(0).toUpperCase() + roleMatch[0].slice(1).toLowerCase()})` : cleanedName;
              break;
            }
          }
        }
        if (!client) {
          for (let line of lines) {
            if (!line.includes("21 CHAMBERS") && !line.includes("YY LEE") && !line.includes("HAVELOCK") && !line.includes("FILE") && !line.includes("SUBJECT")) {
              const possible = cleanString(line).replace(/^[\\[\\]{}|#\\d\\s]+/, '').trim();
              if (/^[A-Z\\s]{4,35}$/i.test(possible) && !/SINGAPORE|ROAD|STREET|AVENUE|BLOCK|TEL/i.test(possible)) {
                client = possible; break;
              }
            }
          }
        }

        let propertyAddress = "-";
        if (dept === "CONVEYANCING") {
          const postalMatch = rawText.match(/(?:Singapore\\s*|\\(?S\\)?\\s*\\(?)(\\d{6})\\)?/i);
          let foundPostal = (postalMatch && postalMatch[1] !== "059763") ? postalMatch[1] : null;

          let addressParts = [];
          for (let line of lines) {
            if (/HAVELOCK|059763|TEL:|FAX:/i.test(line)) continue;
            if (/Decoupling of|Purchase of|Sale of|Block|Blk|Lot No|#\\d+|Fernvale|Close|Street|Road|Avenue|Lane|Drive|Crescent|Way/i.test(line)) {
              addressParts.push(line);
            }
          }

          if (addressParts.length > 0) {
            let rawAddr = addressParts.join(' ');
            propertyAddress = rawAddr
              .replace(/^[0-9\\s|\\[\\]{}Jj§]*/g, '')
              .replace(/(?:SUBJECT|MATTER)\\s*\\|?\\s*/gi, '')
              .replace(/(?:Decoupling|Purchase|Sale|Transfer|Sub-Sale)\\s+of\\s*\\|?\\s*/gi, '')
              .replace(/[|§{}[\\]]/g, ' ')
              .replace(/\\s{2,}/g, ' ')
              .trim();

            if (foundPostal && !propertyAddress.includes(foundPostal)) {
              propertyAddress += `, Singapore ${foundPostal}`;
            }
          }
        } else {
          const courtMatch = rawText.match(/(FC\\/[A-Z0-9\\/]+|HCF\\/[A-Z0-9\\/]+|MSS\\s*\\d+|SS\\s*\\d+)\\s*(\\d+\\/\\d+)?/i);
          if (courtMatch) propertyAddress = courtMatch[0];
        }

        document.getElementById('editLine1').value = fileNo;
        document.getElementById('editLine2').value = client;
        document.getElementById('editLine4').value = propertyAddress;

        const select = document.getElementById('editLine3Select');
        select.value = matchedMatter;
        handleMatterChange(matchedMatter);
      }

      function cleanString(str) {
        return str.replace(/\\+65|\\b\\d{8}\\b|[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}/g, '').replace(/[\\r\\n|]+/g, ' ').replace(/\\s{2,}/g, ' ').trim();
      }

      function openModal(index = -1) {
        currentEditingIndex = index;
        const modal = document.getElementById('editModal');
        modal.style.display = 'flex';

        if (index >= 0) {
          const item = fileQueue[index];
          document.getElementById('editLine1').value = item.line1;
          document.getElementById('editLine2').value = item.line2;
          document.getElementById('editLine4').value = item.line4;

          const select = document.getElementById('editLine3Select');
          if (TAXONOMY_OPTIONS[document.getElementById('departmentSelect').value].includes(item.line3)) {
            select.value = item.line3;
            handleMatterChange(item.line3);
          } else {
            select.value = "OTHER";
            handleMatterChange("OTHER");
            document.getElementById('editLine3Custom').value = item.line3;
          }
          document.getElementById('editLine4').focus();
        } else {
          document.getElementById('editLine1').value = "";
          document.getElementById('editLine2').value = "";
          document.getElementById('editLine4').value = "";
          document.getElementById('editLine3Custom').value = "";
          document.getElementById('editLine3Select').selectedIndex = 0;
          handleMatterChange(document.getElementById('editLine3Select').value);
        }
      }

      function closeModal() { document.getElementById('editModal').style.display = 'none'; }
      function openManualEntry() { openModal(-1); document.getElementById('editLine1').focus(); }

      function saveRecordFromModal() {
        const l1 = document.getElementById('editLine1').value.trim();
        const l2 = document.getElementById('editLine2').value.trim();
        const sel3 = document.getElementById('editLine3Select').value;
        const l3 = (sel3 === "OTHER") ? document.getElementById('editLine3Custom').value.trim() : sel3;
        let l4 = document.getElementById('editLine4').value.trim() || "-";

        if (!l1 || !l2 || !l3) { alert("Please ensure File Number, Client Name, and Matter Type are populated."); return; }

        const record = { line1: l1, line2: l2, line3: l3, line4: l4 };
        if (currentEditingIndex >= 0) fileQueue[currentEditingIndex] = record;
        else fileQueue.push(record);

        renderQueue();
        updateCapacityDisplay();
        saveCartonState();
        closeModal();

        const dept = document.getElementById('departmentSelect').value;
        const { limit } = CAPACITY_LIMITS[dept];
        if (fileQueue.length >= limit) {
          if (confirm(`Carton limit of ${limit} files reached! Export .docx and increment carton?`)) {
            generateWordDoc(true);
          }
        }
      }

      function deleteRecord(index) {
        if (confirm("Remove file record?")) {
          fileQueue.splice(index, 1);
          renderQueue();
          updateCapacityDisplay();
          saveCartonState();
        }
      }

      function renderQueue() {
        const container = document.getElementById('fileQueueContainer');
        container.innerHTML = "";
        document.getElementById('queueCount').textContent = `${fileQueue.length} Files in Active Box`;

        fileQueue.forEach((item, idx) => {
          const div = document.createElement('div');
          div.className = 'file-item';
          div.onclick = (e) => { if (!e.target.classList.contains('del-btn')) openModal(idx); };
          div.innerHTML = `
            <button class="del-btn" onclick="deleteRecord(${idx})">&times;</button>
            <div><span class="line-label">1.</span> <span class="file-no">${item.line1}</span></div>
            <div><span class="line-label">2.</span> ${item.line2}</div>
            <div><span class="line-label">3.</span> <span class="matter-type">${item.line3}</span></div>
            ${item.line4 && item.line4 !== "-" ? `<div><span class="line-label">4.</span> ${item.line4}</div>` : ""}
          `;
          container.appendChild(div);
        });
      }

      async function generateWordDoc(autoIncrement = false) {
        if (fileQueue.length === 0) { alert("Queue is empty. Scan files first."); return; }
        const cartonInput = document.getElementById('cartonNumberInput');
        const cartonNo = cartonInput.value.trim() || "YYLEE-BOX-RECORD";
        const { Document, Packer, Paragraph, Table, TableRow, TableCell, WidthType, AlignmentType, TextRun, BorderStyle } = docx;

        const thinBorder = {
          top: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
          bottom: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
          left: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
          right: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
        };

        const tableRows = [
          new TableRow({
            children: [
              new TableCell({
                width: { size: 10, type: WidthType.PERCENTAGE }, borders: thinBorder,
                children: [new Paragraph({ text: "TO : CROWN WORLDWIDE PTE LTD" }), new Paragraph({ children: [new TextRun({ text: "Carton No.\\n" + cartonNo, bold: true })] })]
              }),
              new TableCell({
                width: { size: 55, type: WidthType.PERCENTAGE }, borders: thinBorder,
                children: [new Paragraph({ children: [new TextRun({ text: "CARTON DETAILS", bold: true })] }), new Paragraph({ text: "1. File Number\\n2. Client\\n3. Matter Type\\n4. Case No. / Property Address" })]
              }),
              new TableCell({
                width: { size: 18, type: WidthType.PERCENTAGE }, borders: thinBorder,
                children: [new Paragraph({ children: [new TextRun({ text: "EXPECTED DESTRUCTION DATE", bold: true })] })]
              }),
              new TableCell({
                width: { size: 17, type: WidthType.PERCENTAGE }, borders: thinBorder,
                children: [new Paragraph({ children: [new TextRun({ text: "FOR OFFICIAL USE", bold: true })] })]
              })
            ]
          })
        ];

        fileQueue.forEach((item, index) => {
          const pList = [
            new Paragraph({ children: [new TextRun({ text: "1. " + item.line1, bold: true })] }),
            new Paragraph({ text: "2. " + item.line2 }),
            new Paragraph({ text: "3. " + item.line3 }),
          ];
          if (item.line4 && item.line4 !== "-") pList.push(new Paragraph({ text: "4. " + item.line4 }));

          tableRows.push(
            new TableRow({
              children: [
                new TableCell({ borders: thinBorder, children: [new Paragraph({ text: `${index + 1}.`, alignment: AlignmentType.CENTER })] }),
                new TableCell({ borders: thinBorder, children: pList }),
                new TableCell({ borders: thinBorder, children: [new Paragraph("")] }),
                new TableCell({ borders: thinBorder, children: [new Paragraph("")] })
              ]
            })
          );
        });

        const doc = new Document({
          styles: { default: { document: { run: { font: "Arial", size: 18 } } } },
          sections: [{
            properties: { page: { margin: { top: 720, right: 720, bottom: 720, left: 720 } } },
            children: [new Table({ rows: tableRows, width: { size: 100, type: WidthType.PERCENTAGE } })]
          }]
        });

        const blob = await Packer.toBlob(doc);
        saveAs(blob, `Crown Box Record (${cartonNo}).docx`);

        if (autoIncrement) {
          fileQueue = [];
          renderQueue();
          localStorage.removeItem('chambersos_active_queue');
          const numMatch = cartonNo.match(/\\d+/g);
          if (numMatch) {
            const lastNum = numMatch[numMatch.length - 1];
            const nextNum = parseInt(lastNum, 10) + 1;
            cartonInput.value = cartonNo.replace(new RegExp(lastNum + "(?!.*" + lastNum + ")"), nextNum);
            saveCartonState();
          }
          updateCapacityDisplay();
        }
      }
    </script>
    </body>
    </html>
    """

    components.html(crown_scanner_component, height=1050, scrolling=True)

# ==============================================================================
# WING 3: CROWN ARCHIVAL LOCATOR & SEARCH ENGINE
# ==============================================================================
with tab_locator:
    st.markdown("### 🔍 Crown Worldwide Warehouse Archive Locator")
    st.caption(
        "Instant multi-index search across 8,000+ historical firm files, physical cartons, and client jackets."
    )

    db_path = "crown_base.db"

    if not os.path.exists(db_path):
        st.error(
            "⚠️ `crown_base.db` not found in repository root. Place your compiled database file in the project folder."
        )
    else:
        # Top Metrics Bar & Database Export Tool (Cloud Safeguard)
        col_metrics, col_download = st.columns([3, 1])

        try:
            conn = sqlite3.connect(db_path, timeout=10)
            total_records = conn.execute(
                "SELECT COUNT(*) FROM archive_records"
            ).fetchone()[0]
            total_cartons = conn.execute(
                "SELECT COUNT(DISTINCT carton_no) FROM archive_records"
            ).fetchone()[0]
            conn.close()

            with col_metrics:
                cm1, cm2, cm3 = st.columns(3)
                with cm1:
                    st.metric("Total Indexed Matters", f"{total_records:,}")
                with cm2:
                    st.metric("Total Physical Cartons", f"{total_cartons:,}")
                with cm3:
                    st.metric("Database Engine", "SQLite B-Tree")

            with col_download:
                st.write("")
                with open(db_path, "rb") as db_file:
                    st.download_button(
                        label="📥 Download Master DB",
                        data=db_file,
                        file_name="crown_base.db",
                        mime="application/x-sqlite3",
                        help="Download the updated SQLite database file to commit back to Git after batch scanning.",
                        use_container_width=True,
                    )
        except Exception as e:
            st.warning(f"Metadata read error: {e}")

        # Ingestion Chute for Newly Exported Cartons
        with st.expander("📥 Ingest Newly Exported Carton (.docx)", expanded=False):
            st.caption("Drop any newly exported Crown Box Record (`.docx`) here to permanently index its files into the search catalog with duplicate protection.")
            new_box_file = st.file_uploader("Upload exported carton .docx", type=["docx"], key="ingest_box_uploader")
            if new_box_file:
                if st.button("Slide Carton into Archive Catalog", type="primary"):
                    try:
                        doc = docx.Document(new_box_file)
                        c_no = "UNKNOWN"
                        if doc.tables and len(doc.tables[0].rows) > 0:
                            header_cell = doc.tables[0].rows[0].cells[0].text
                            c_match = re.search(r"Carton No\.?\s*([\s\S]+?)(?=CARTON DETAILS|$)", header_cell, re.IGNORECASE)
                            c_no = re.sub(r"\s+", " ", c_match.group(1)).strip() if c_match else new_box_file.name.replace(".docx", "")
                        
                        inserted = 0
                        duplicates = 0
                        conn = sqlite3.connect(db_path, timeout=10)
                        cur = conn.cursor()
                        
                        # Ensure unique constraint exists to block duplicates safely
                        cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_unique_carton_file ON archive_records(carton_no, file_no)")

                        for t in doc.tables:
                            for row in t.rows[1:]:
                                if len(row.cells) >= 2:
                                    txt = row.cells[1].text.strip()
                                    lines = [l.strip() for l in txt.split("\n") if l.strip()]
                                    f_no, cl, mt, mt_data = "-", "-", "-", "-"
                                    for line in lines:
                                        if line.startswith("1."): f_no = re.sub(r"^1\.?\s*", "", line).strip()
                                        elif line.startswith("2."): cl = re.sub(r"^2\.?\s*", "", line).strip()
                                        elif line.startswith("3."): mt = re.sub(r"^3\.?\s*", "", line).strip()
                                        elif line.startswith("4."): mt_data = re.sub(r"^4\.?\s*", "", line).strip()
                                    if f_no != "-" or cl != "-":
                                        try:
                                            cur.execute(
                                                """INSERT INTO archive_records (carton_no, file_no, client_name, matter_type, target_metadata, source_file)
                                                   VALUES (?, ?, ?, ?, ?, ?)""",
                                                (c_no, f_no, cl, mt, mt_data, new_box_file.name)
                                            )
                                            inserted += 1
                                        except sqlite3.IntegrityError:
                                            duplicates += 1

                        conn.commit()
                        conn.close()
                        st.success(f"Successfully indexed {inserted} matters from {c_no}! ({duplicates} duplicates skipped)")
                        st.rerun()
                    except Exception as err:
                        st.error(f"Ingestion failed: {err}")

        st.markdown("---")

        # Search Bar & Filter Controls
        search_col, filter_col = st.columns([3, 1])
        with search_col:
            search_query = st.text_input(
                "Search Archives",
                placeholder="Search by Client Name, File No (e.g. 201900782), Property Address, or Postal Code...",
                key="archive_search_input",
            )
        with filter_col:
            field_filter = st.selectbox(
                "Filter Field",
                ["All Fields", "Client Name", "File Number", "Address / Case No"],
            )

        if search_query.strip():
            q = f"%{search_query.strip()}%"
            conn = sqlite3.connect(db_path, timeout=10)

            if field_filter == "Client Name":
                sql = """SELECT carton_no, file_no, client_name, matter_type, target_metadata, source_file 
                         FROM archive_records WHERE client_name LIKE ? ORDER BY id DESC LIMIT 100"""
                params = (q,)
            elif field_filter == "File Number":
                sql = """SELECT carton_no, file_no, client_name, matter_type, target_metadata, source_file 
                         FROM archive_records WHERE file_no LIKE ? ORDER BY id DESC LIMIT 100"""
                params = (q,)
            elif field_filter == "Address / Case No":
                sql = """SELECT carton_no, file_no, client_name, matter_type, target_metadata, source_file 
                         FROM archive_records WHERE target_metadata LIKE ? ORDER BY id DESC LIMIT 100"""
                params = (q,)
            else:
                sql = """SELECT carton_no, file_no, client_name, matter_type, target_metadata, source_file 
                         FROM archive_records 
                         WHERE client_name LIKE ? OR file_no LIKE ? OR target_metadata LIKE ? OR carton_no LIKE ?
                         ORDER BY id DESC LIMIT 100"""
                params = (q, q, q, q)

            results = conn.execute(sql, params).fetchall()
            conn.close()

            if results:
                st.success(f"Found **{len(results)}** matching record(s):")

                for r in results:
                    carton, f_no, client, matter, meta, src = r

                    with st.container():
                        st.markdown(
                            f"""
                            <div style="background:#ffffff; border:1px solid #e2e8f0; border-left:4px solid #1e40af; border-radius:8px; padding:12px 16px; margin-bottom:10px; box-shadow:0 1px 2px rgba(0,0,0,0.03);">
                                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
                                    <span style="font-size:1.05rem; font-weight:700; color:#1e40af;">File: {f_no}</span>
                                    <span style="background:#eff6ff; color:#1e40af; border:1px solid #bfdbfe; font-size:0.8rem; font-weight:700; padding:3px 10px; border-radius:6px;">📦 Carton: {carton}</span>
                                </div>
                                <div style="font-size:0.95rem; font-weight:600; color:#0f172a; margin-bottom:4px;">Client: {client}</div>
                                <div style="font-size:0.85rem; color:#059669; font-weight:600; margin-bottom:4px;">Matter: {matter}</div>
                                {f'<div style="font-size:0.85rem; color:#475569;">📍 Property / Case: {meta}</div>' if meta and meta != '-' else ''}
                                <div style="font-size:0.75rem; color:#94a3b8; margin-top:6px;">Source: {src}</div>
                            </div>
                            """,
                            unsafe_allow_html=True,
                        )
            else:
                st.info(f"No records found matching '{search_query.strip()}'.")
