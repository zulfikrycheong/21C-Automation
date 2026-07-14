import streamlit as st
import docx
import re
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
import os
import base64
import json
import io

# ReportLab layout framework components
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# Force clean full-width canvas layout workspace view
st.set_page_config(page_title="21 Chambers Client List", layout="wide")

if os.path.exists("Company Logo.png"):
    st.image("Company Logo.png", width=250)

st.title("📂 21 Chambers Automator")
st.markdown("##### *Dual-Stream Intake Pipeline & Native Vector PDF Center*")
st.write("Drag and drop your document below to synchronize the master matrix and download print-ready layouts.")
st.markdown("---")

# --- 1. GOOGLE SHEETS MATRIX SYNC ---
GOOGLE_SHEET_NAME = "Lazy Automation"  
SHEET_TAB_NAME = datetime.now().strftime("%B %Y")
CURRENT_YEAR = datetime.now().strftime("%Y")

def get_google_sheet():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    if os.path.exists("credentials.json"):
        creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
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
            source_sheet_id=template_sheet.id, new_sheet_name=SHEET_TAB_NAME, insert_sheet_index=0 
        )
    return sheet

# --- 2. VECTOR PDF ENGINE ---
def generate_perfect_pdf(matter_no, clients_text, contacts_text, matter_type, date_opened):
    buffer = io.BytesIO()
    
    # Page Setup: A4 with exact 0.295" left/right margins (21.24 points)
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        leftMargin=21.24, rightMargin=21.24,
        topMargin=36.0, bottomMargin=36.0
    )
    
    # Custom Typography Style Profiles matching your 20pt Times New Roman spec sheet exactly
    style_normal_20 = ParagraphStyle('Norm20', fontName='Times-Roman', fontSize=20, leading=26, alignment=TA_LEFT)
    style_bold_20 = ParagraphStyle('Bold20', fontName='Times-Bold', fontSize=20, leading=26, alignment=TA_LEFT)
    style_firm_title = ParagraphStyle('FirmTitle', fontName='Times-Bold', fontSize=20, leading=26, alignment=TA_CENTER)
    style_firm_body = ParagraphStyle('FirmBody', fontName='Times-Roman', fontSize=20, leading=26, alignment=TA_CENTER)
    
    # The Mega Footer Code (Arial Bold 109pt)
    style_giant_foot = ParagraphStyle('GiantFoot', fontName='Helvetica-Bold', fontSize=109, leading=115, alignment=TA_CENTER)
    
    story = []
    
    # ----------------------------------------------------
    # TOP SECTION: Spacing & Party Layout (REPLICATED WITH BOTTOM CAGE RULES)
    # ----------------------------------------------------
    client_lines = [line.strip() for line in clients_text.split('\n') if line.strip()]
    contact_lines = [line.strip() for line in contacts_text.split('\n') if line.strip()]
    
    party_elements = []
    
    # Process Applicant Block
    if len(client_lines) > 0:
        party_elements.append(Paragraph(client_lines[0], style_normal_20))
    if len(contact_lines) > 0:
        party_elements.append(Paragraph(contact_lines[0], style_normal_20))
        
    party_elements.append(Spacer(1, 14)) # Clean vertical separation gap
    
    # Process Respondent Block
    if len(client_lines) > 1:
        party_elements.append(Paragraph(client_lines[1], style_normal_20))
    if len(contact_lines) > 1:
        party_elements.append(Paragraph(contact_lines[1], style_normal_20))
        
    # Columns: Left Spacer (1.333") vs Right Data Column (6.346")
    col_w_left = 1.333 * 72
    col_w_right = 6.346 * 72
    
    # FIX: Replicated the exact geometric grid cage parameters onto the top container structure
    top_table = Table([["", party_elements]], colWidths=[col_w_left, col_w_right])
    top_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 1.5, colors.black),  # 🏛️ Cage lines explicitly added here!
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING', (0,0), (-1,-1), 7.2),         # Exact 100 dxa spacing
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(top_table)
    story.append(Spacer(1, 24))
    
    # ----------------------------------------------------
    # CENTRAL HEADER BLOCK (Untouched, Perfect Center)
    # ----------------------------------------------------
    story.append(Paragraph("21 CHAMBERS LLC", style_firm_title))
    story.append(Spacer(1, 4))
    story.append(Paragraph("2 HAVELOCK ROAD #06-17<br/>HAVELOCK 2<br/>SINGAPORE 059763", style_firm_body))
    story.append(Spacer(1, 4))
    story.append(Paragraph("TEL: 6224 1848 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; FAX: 6223 3092", style_firm_body))
    story.append(Spacer(1, 40))
    
    # ----------------------------------------------------
    # BOTTOM MATRIX: Auto-Expanding Vector Cage (Untouched)
    # ----------------------------------------------------
    full_matter_name = "Uncontested Divorce"
    if matter_type == "CD": full_matter_name = "Contested Divorce"
    elif matter_type == "Annulment": full_matter_name = "Annulment"
    elif matter_type == "Variation": full_matter_name = "Variation"
    elif matter_type == "Others": full_matter_name = "Others"

    file_block_text = f"{matter_no}<br/>Opening date: {date_opened}<br/>Closure date:"
    
    matrix_rows = [
        [Paragraph("SUBJECT<br/>MATTER", style_bold_20), Paragraph(full_matter_name, style_normal_20)],
        [Paragraph("FILE", style_normal_20), Paragraph(file_block_text, style_normal_20)],
        [Paragraph("Legal Fee", style_normal_20), Paragraph("CASH", style_normal_20)],
        [Paragraph("Remarks", style_normal_20), Paragraph("", style_normal_20)]
    ]
    
    b_col1 = 1.596 * 72
    b_col2 = 6.083 * 72
    
    bottom_table = Table(matrix_rows, colWidths=[b_col1, b_col2])
    bottom_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 1.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING', (0,0), (-1,-1), 7.2),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(bottom_table)
    story.append(Spacer(1, 54)) 
    
    # ----------------------------------------------------
    # BOTTOM SECTION: The Giant Footer Display (Untouched)
    # ----------------------------------------------------
    story.append(Paragraph(matter_no, style_giant_foot))
    
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

# --- 3. DOCX PARSING INTAKE ENGINE ---
def extract_matter_data(doc_path):
    doc = docx.Document(doc_path)
    text_lines = []
    for para in doc.paragraphs:
        if para.text.strip(): text_lines.append(para.text)
    for table in doc.tables:
        for row in table.rows:
            row_text = [cell.text.strip() for cell in row.cells if cell.text.strip()]
            if row_text: text_lines.append(" ".join(row_text))
                
    full_text = "\n".join(text_lines)
    
    # ⚖️ Matter Type Assignment
    if "uncontested divorce" in full_text.lower(): matter_type = "UD"
    elif "contested divorce" in full_text.lower(): matter_type = "CD"
    elif "annulment" in full_text.lower(): matter_type = "Annulment"
    elif "variation" in full_text.lower(): matter_type = "Variation"
    else: matter_type = "Others"
    
    # 🏢 1. Raw Text Block Extraction (Targeting text sections exactly)
    # Splitting the document loosely by sections to isolate Applicant vs Respondent info
    app_section = ""
    res_section = ""
    
    # Find positions to split text logically if both headers exist
    app_idx = full_text.lower().find("applicant")
    res_idx = full_text.lower().find("respondent")
    
    if app_idx != -1 and res_idx != -1:
        if app_idx < res_idx:
            app_section = full_text[app_idx:res_idx]
            res_section = full_text[res_idx:]
        else:
            res_section = full_text[res_idx:app_idx]
            app_section = full_text[app_idx:]
    elif app_idx != -1:
        app_section = full_text[app_idx:]
    elif res_idx != -1:
        res_section = full_text[res_idx:]

    # --- APPLICANT PARSING NODE ---
    app_match = re.search(r"Applicant\s*–\s*([^\n\d]+)", full_text, re.IGNORECASE)
    if app_match:
        clean_app = re.split(r"\s*[-\s–]\s*upload", app_match.group(1), flags=re.IGNORECASE)[0]
        applicant_name = clean_app.strip().upper()
    else:
        applicant_name = "NIL" # Fallback: No name found

    # Scan specifically inside the applicant's text region for their contacts
    app_target_text = app_section if app_section else full_text
    app_target_text = app_target_text.replace("6224 1848", "").replace("6223 3092", "") # Exclude firm lines
    
    app_mobs = re.findall(r"\+?\b(?:65|60|1|44)?[ \-]?[89]\d{3}[ \-]?\d{4}\b|\+?60[ \-]?1\d[ \-]?\d{2,3}[ \-]?\d{4}\b|\b01\d[ \-]?\d{2,3}[ \-]?\d{4}\b", app_target_text)
    app_emails = re.findall(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", app_target_text)
    
    app_mob = "NIL"
    if app_mobs:
        clean_mob = re.sub(r"[\s\-+]", "", app_mobs[0])
        if clean_mob.startswith('01'): clean_mob = f"60{clean_mob[1:]}"
        
        if clean_mob.startswith(('8', '9')) and len(clean_mob) == 8:
            app_mob = f"+65 {clean_mob}"
        else:
            app_mob = f"+{clean_mob[:2]} {clean_mob[2:]}"
            
    app_email = app_emails[0] if app_emails else "NIL" # Fallback: Put NIL if no email exists

    # --- RESPONDENT PARSING NODE ---
    res_match = re.search(r"Respondent\s*–\s*([^\n\d]+)", full_text, re.IGNORECASE)
    if res_match:
        clean_res = re.split(r"\s*[-\s–]\s*upload", res_match.group(1), flags=re.IGNORECASE)[0]
        respondent_name = clean_res.strip().upper()
    else:
        respondent_name = "NIL" # Fallback: No name found

    # Scan specifically inside the respondent's text region for their contacts
    res_target_text = res_section if res_section else full_text
    res_target_text = res_target_text.replace("6224 1848", "").replace("6223 3092", "") # Exclude firm lines
    
    res_mobs = re.findall(r"\+?\b(?:65|60|1|44)?[ \-]?[89]\d{3}[ \-]?\d{4}\b|\+?60[ \-]?1\d[ \-]?\d{2,3}[ \-]?\d{4}\b|\b01\d[ \-]?\d{2,3}[ \-]?\d{4}\b", res_target_text)
    res_emails = re.findall(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", res_target_text)
    
    res_mob = "NIL"
    if res_mobs:
        clean_mob = re.sub(r"[\s\-+]", "", res_mobs[0])
        if clean_mob.startswith('01'): clean_mob = f"60{clean_mob[1:]}"
        
        if clean_mob.startswith(('8', '9')) and len(clean_mob) == 8:
            res_mob = f"+65 {clean_mob}"
        else:
            res_mob = f"+{clean_mob[:2]} {clean_mob[2:]}"
            
    res_email = res_emails[0] if res_emails else "NIL" # Fallback: Put NIL if no email exists

    # --- RE-MAPPING THE DUAL ARRAY OUTPUTS ---
    # We cleanly arrange them into the text variables exactly how our cover sheet template prints them
    clients_field = f"APPLICANT - {applicant_name}\nRESPONDENT - {respondent_name}"
    
    # We compile the contact lines dynamically. If a line has a phone or email, it prints it.
    # Otherwise, it cleanly outputs the mapped fallback tokens.
    contacts_field = f"{app_mob} {app_email}\n{res_mob} {res_email}".strip()
    
    # 🔹 Referral tracking (established logic remains completely untouched)
    referral = "Google"
    if "referral" in full_text.lower():
        ref_match = re.search(r"referral\s*[\s,:\-–]\s*(\w+)", full_text, re.IGNORECASE)
        if ref_match and "jav" in ref_match.group(1).lower(): referral = "Javern"
    elif "jav" in full_text.lower(): referral = "Javern"
            
    return matter_type, clients_field, contacts_field, referral

# --- 4. STREAMLIT FRAMEWORK FLOWS ---
if "uploader_key" not in st.session_state: st.session_state["uploader_key"] = 0
if "previous_files" not in st.session_state: st.session_state["previous_files"] = []
if "pdf_binary_store" not in st.session_state: st.session_state["pdf_binary_store"] = {}

# Fixed CSS: Massively scales the drop zone while respecting Dark Mode typography
st.markdown("""
    <style>
        /* Expands the target footprint area */
        [data-testid="stFileUploaderDropzone"] {
            padding: 6rem 3rem !important;
            border-radius: 12px !important;
            background-color: #1e222b !important; /* Dark sleek widget background */
            border: 2px dashed #4b5563 !important; /* Defined tracking border */
        }
        
        /* Forces the text labels to pop out in clean white/light gray */
        [data-testid="stFileUploaderDropzone"] label, 
        [data-testid="stFileUploaderDropzone"] div {
            color: #ffffff !important;
        }
    </style>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "Drag and drop Open File Sheets (.docx) here", 
    type=["docx"], accept_multiple_files=True, key=f"uploader_{st.session_state['uploader_key']}"
)

current_file_names = [f.name for f in uploaded_files] if uploaded_files else []
if current_file_names != st.session_state["previous_files"]:
    st.session_state["previous_files"] = current_file_names
    st.session_state["pdf_binary_store"] = {}

if uploaded_files:
    st.markdown("---")
    st.subheader("🖨️ Production Queue Confirmation")
    
    c1, c2 = st.columns(2)
    with c1:
        if st.button("✅ YES - Execute Streams & Generate PDFs", use_container_width=True, type="primary"):
            if not st.session_state["pdf_binary_store"]:
                with st.spinner("⚡ Activating dual architectural cages..."):
                    sheet = get_google_sheet()
                    
                    try:
                        vals = sheet.col_values(3)[1:]
                        valid = [int(v.strip()) for v in vals if v.strip().isdigit()]
                    except:
                        valid = []
                    current_max = max(valid) if valid else 20260728
                    
                    temp_store = {}
                    for f in uploaded_files:
                        matter_nos = sheet.col_values(3)
                        mx, mx_idx = -1, 1
                        for idx, val in enumerate(matter_nos):
                            if str(val).strip().isdigit() and int(str(val).strip()) > mx:
                                mx = int(str(val).strip())
                                mx_idx = idx + 1
                        
                        target_row = mx_idx + 1 if mx != -1 else 2
                        if mx != -1: current_max = mx
                        
                        current_max += 1
                        new_no = str(current_max)
                        next_idx = target_row - 1
                        today_str = datetime.now().strftime("%d %B %Y").lstrip("0").upper()
                        
                        m_type, cls, cnt, ref = extract_matter_data(f)
                        
                        new_row = [next_idx, today_str, new_no, m_type, cls, cnt, ref, "Yes", ""]
                        sheet.update(range_name=f"A{target_row}:I{target_row}", values=[new_row])
                        
                        # Process layout engine output
                        pdf_output_bytes = generate_perfect_pdf(new_no, cls, cnt, m_type, today_str)
                        temp_store[f.name] = (new_no, pdf_output_bytes)
                        st.toast(f"Synchronized Matrix: Matter {new_no}", icon="🔹")
                        
                    st.session_state["pdf_binary_store"] = temp_store
                    st.balloons()
                    
    with c2:
        if st.button("❌ NO - Abort Operational Batch", use_container_width=True):
            st.session_state["uploader_key"] += 1
            st.session_state["previous_files"] = []
            st.session_state["pdf_binary_store"] = {}
            st.rerun()

if st.session_state["pdf_binary_store"]:
    st.markdown("---")
    st.success("🎉 **Data routing completely finalized. Download the printable cover sheet below:**")
    for fname, (m_no, pdf_bytes) in st.session_state["pdf_binary_store"].items():
        st.download_button(
            label=f"🖨️ Download Production Cover Sheet PDF (Matter File Number: {m_no})",
            data=pdf_bytes,
            file_name=f"21Chambers_Cover_{m_no}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
