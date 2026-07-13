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

# ReportLab core engine components for high-fidelity vector PDF generation
from reportlab.lib.pagesizes import a4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

st.set_page_config(page_title="21 Chambers Client List", layout="centered")

# --- UI INTERFACE GRAPHICS & LOGO ---
if os.path.exists("Company Logo.png"):
    st.image("Company Logo.png", width=250)

st.title("📂 21 Chambers Automator")
st.markdown("##### *Dual-Stream Intake Pipeline & PDF Cover Sheet Generator*")
st.write("Drag and drop your document below to synchronize the master matrix and prepare print layouts.")
st.markdown("---")

# --- 1. GOOGLE SHEETS AUTHENTICATION LINK ---
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
        duplicated_sheet = workbook.duplicate_sheet(
            source_sheet_id=template_sheet.id, new_sheet_name=SHEET_TAB_NAME, insert_sheet_index=0 
        )
        sheet = duplicated_sheet
    return sheet

# --- 2. VECTOR PDF BLUEPRINT GENERATOR ---
def generate_cover_pdf(matter_no, clients_text, contacts_text, matter_type, date_opened):
    buffer = io.BytesIO()
    
    # Establish dynamic A4 canvas profile with strict 0.8-inch operational margins
    doc = SimpleDocTemplate(
        buffer, pagesize=a4,
        rightMargin=57.6, leftMargin=57.6, topMargin=57.6, bottomMargin=57.6
    )
    
    styles = getSampleStyleSheet()
    
    # Custom high-fidelity typography mappings matching your specific physical sheet profile
    style_top_box = ParagraphStyle('TopBox', fontName='Helvetica', fontSize=13, leading=16)
    style_top_box_bold = ParagraphStyle('TopBoxBold', fontName='Helvetica-Bold', fontSize=13, leading=16)
    
    style_firm_title = ParagraphStyle('FirmTitle', fontName='Helvetica-Bold', fontSize=20, leading=24, alignment=TA_CENTER)
    style_firm_body = ParagraphStyle('FirmBody', fontName='Helvetica', fontSize=13, leading=16, alignment=TA_CENTER)
    
    style_matrix_lbl = ParagraphStyle('MatrixLbl', fontName='Helvetica-Bold', fontSize=13, leading=16)
    style_matrix_val = ParagraphStyle('MatrixVal', fontName='Helvetica', fontSize=13, leading=16)
    
    style_giant_foot = ParagraphStyle('GiantFoot', fontName='Helvetica-Bold', fontSize=72, leading=76, alignment=TA_CENTER)
    
    story = []
    
    # A. TOP CONTAINER BOX BORDER REPLICATION
    client_lines = [line.strip() for line in clients_text.split('\n') if line.strip()]
    contact_lines = [line.strip() for line in contacts_text.split('\n') if line.strip()]
    
    top_box_elements = []
    for line in client_lines:
        current_style = style_top_box_bold if ("APPLICANT" in line or "RESPONDENT" in line) else style_top_box
        top_box_elements.append(Paragraph(line, current_style))
    for line in contact_lines:
        top_box_elements.append(Paragraph(line, style_top_box))
        
    # Standard printable width for accurate single-cell bounding mapping
    printable_width = a4[0] - 115.2 
    
    top_table = Table([[top_box_elements]], colWidths=[printable_width])
    top_table.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 1.5, colors.black),
        ('PADDING', (0,0), (-1,-1), 12),
        ('BOTTOMPADDING', (0,0), (-1,-1), 12),
    ]))
    story.append(top_table)
    story.append(Spacer(1, 28))
    
    # B. CENTRAL FIRM SIGNATURE BLOCK
    story.append(Paragraph("21 CHAMBERS LLC", style_firm_title))
    story.append(Spacer(1, 4))
    story.append(Paragraph("2 HAVELOCK ROAD #06-17<br/>HAVELOCK 2<br/>SINGAPORE 059763", style_firm_body))
    story.append(Spacer(1, 4))
    story.append(Paragraph("TEL: 6224 1848 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; FAX: 6223 3092", style_firm_body))
    story.append(Spacer(1, 36))
    
    # C. CORE METADATA MATRIX TABLE
    subject_label = "GLA for Estate of ..."
    if client_lines:
        clean_name = client_lines[0].replace("APPLICANT - ", "")
        subject_label = f"{matter_type} for {clean_name}"
        
    file_block_text = f"<b>{matter_no}</b><br/>Opening date: {date_opened}<br/>Closure date:"
    
    matrix_data = [
        [Paragraph("SUBJECT MATTER", style_matrix_lbl), Paragraph(subject_label, style_matrix_val)],
        [Paragraph("FILE", style_matrix_lbl), Paragraph(file_block_text, style_matrix_val)],
        [Paragraph("Legal Fee", style_matrix_lbl), Paragraph("CASH", style_matrix_val)],
        [Paragraph("Remarks", style_matrix_lbl), Paragraph("", style_matrix_val)]
    ]
    
    matrix_table = Table(matrix_data, colWidths=[printable_width * 0.25, printable_width * 0.75])
    matrix_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 1.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('PADDING', (0,0), (-1,-1), 10),
        ('BOTTOMPADDING', (0,0), (-1,-1), 10),
    ]))
    story.append(matrix_table)
    story.append(Spacer(1, 40))
    
    # D. GIANT FOOTER CODE TRACKING DISPLAY
    story.append(Paragraph(matter_no, style_giant_foot))
    
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

# --- 3. RAW INPUT DOCX SCRAPING ENGINE ---
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
    
    if "uncontested divorce" in full_text.lower(): matter_type = "UD"
    elif "contested divorce" in full_text.lower(): matter_type = "CD"
    elif "annulment" in full_text.lower(): matter_type = "Annulment"
    elif "variation" in full_text.lower(): matter_type = "Variation"
    else: matter_type = "Others"
    
    app_match = re.search(r"Applicant\s*–\s*([^,\n\d]+)", full_text, re.IGNORECASE)
    res_match = re.search(r"Respondent\s*–\s*([^–,\n\d]+)", full_text, re.IGNORECASE)
    
    applicant = "APPLICANT - NIL"
    if app_match:
        clean_app = re.split(r"\s*[-\s–]\s*upload", app_match.group(1), flags=re.IGNORECASE)[0]
        applicant = f"APPLICANT - {clean_app.strip().upper()}"
        
    respondent = "RESPONDENT - NIL"
    if res_match:
        clean_res = re.split(r"\s*[-\s–]\s*upload", res_match.group(1), flags=re.IGNORECASE)[0]
        respondent = f"RESPONDENT - {clean_res.strip().upper()}"
        
    clients_field = f"{applicant}\n{respondent}"
    
    raw_mobiles = re.findall(r"\b[89]\d{3}[ \-]?\d{4}\b", full_text)
    mobiles = [re.sub(r"[\s\-]", "", mob) for mob in raw_mobiles]
    emails = re.findall(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", full_text)
    
    contacts = []
    for i in range(max(len(mobiles), len(emails))):
        mob = f"+65 {mobiles[i]}" if i < len(mobiles) else ""
        em = emails[i] if i < len(emails) else ""
        contacts.append(f"{mob} {em}".strip())
    contacts_field = "\n".join(contacts)
    
    referral = "Google"
    if "referral" in full_text.lower():
        ref_match = re.search(r"referral\s*[\s,:\-–]\s*(\w+)", full_text, re.IGNORECASE)
        if ref_match and "jav" in ref_match.group(1).lower(): referral = "Javern"
    elif "jav" in full_text.lower(): referral = "Javern"
            
    return matter_type, clients_field, contacts_field, referral

# --- SIDEBAR ACCESS PANEL CONTROL ---
with st.sidebar:
    st.image("Company Logo.png", use_container_width=True)  
    st.markdown("### 🛠️ Operation Logistics")
    st.info("Dual-Stream Automation: Syncing master sheet records and vector PDF printers concurrently.")
    st.markdown("---")
    st.markdown("### 📋 System Access Email")
    try:
        if os.path.exists("credentials.json"):
            with open("credentials.json", "r") as f:
                sys_email = json.load(f).get("client_email", "Not found")
        else:
            creds_dict = json.loads(base64.b64decode(st.secrets["encoded_creds"]))
            sys_email = creds_dict.get("client_email", "Not found")
        st.code(sys_email, language="text")  
    except Exception:
        st.error("Access configuration unresolved.")

st.markdown("---")

# --- 4. DATA PROCESSING AND INTERACTIVE RUNTIME FLOW ---
if "uploader_key" not in st.session_state: st.session_state["uploader_key"] = 0
if "previous_files" not in st.session_state: st.session_state["previous_files"] = []
if "pdf_data_store" not in st.session_state: st.session_state["pdf_data_store"] = {}

uploaded_files = st.file_uploader(
    "Drag and drop Open File Sheets (.docx) here", 
    type=["docx"], accept_multiple_files=True, key=f"uploader_{st.session_state['uploader_key']}"
)

# Clear tracking states if user clears out files manually
current_file_names = [f.name for f in uploaded_files] if uploaded_files else []
if current_file_names != st.session_state["previous_files"]:
    st.session_state["previous_files"] = current_file_names
    st.session_state["pdf_data_store"] = {}

if uploaded_files:
    st.markdown("---")
    st.subheader("🖨️ Automated Document Routing Center")
    st.write("Documents staged. Would you like to initialize the dual-stream automation pipelines?")
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("✅ YES - Execute Pipelines & Generate PDFs", use_container_width=True, type="primary"):
            if not st.session_state["pdf_data_store"]:
                with st.spinner("⚡ Initializing dual outputs: Syncing cloud matrix & drawing vectors..."):
                    sheet = get_google_sheet()
                    
                    try:
                        matter_nos = sheet.col_values(3)[1:]  
                        valid_numbers = [int(val.strip()) for val in matter_nos if str(val).strip().isdigit()]
                    except Exception:
                        valid_numbers = []
                    
                    if not valid_numbers:
                        if datetime.now().month == 1:
                            current_max_matter = int(f"{CURRENT_YEAR}0000")
                        else:
                            try:
                                from datetime import timedelta
                                prev_sheet = client.open(GOOGLE_SHEET_NAME).worksheet((datetime.now().replace(day=1) - timedelta(days=15)).strftime("%B %Y"))
                                valid_numbers = [int(val.strip()) for val in prev_sheet.col_values(3)[1:] if str(val).strip().isdigit()]
                            except Exception:
                                valid_numbers = [20260728] 
                    
                    current_max_matter = max(valid_numbers) if valid_numbers else 20260728
                    
                    temp_store = {}
                    # Stream 1: Synchronize rows directly to the Google Sheet
                    for doc_file in uploaded_files:
                        matter_nos = sheet.col_values(3) 
                        max_num = -1
                        max_row_index = 1 
                        for idx, val in enumerate(matter_nos):
                            if str(val).strip().isdigit() and int(str(val).strip()) > max_num:
                                max_num = int(str(val).strip())
                                max_row_index = idx + 1 
                        
                        target_row = max_row_index + 1 if max_num != -1 else 2
                        if max_num != -1: current_max_matter = max_num
                        
                        current_max_matter += 1
                        new_matter_no = str(current_max_matter)
                        next_index = target_row - 1
                        today_date = datetime.now().strftime("%d %B %Y").lstrip("0")
                        
                        matter_type, clients, contacts, referral = extract_matter_data(doc_file)
                        
                        new_row = [next_index, today_date, new_matter_no, matter_type, clients, contacts, referral, "Yes", ""]
                        sheet.update(range_name=f"A{target_row}:I{target_row}", values=[new_row])
                        
                        # Stream 2: Render print-ready layout into background binary buffers concurrently
                        pdf_bytes = generate_cover_pdf(new_matter_no, clients, contacts, matter_type, today_date)
                        temp_store[doc_file.name] = (new_matter_no, pdf_bytes)
                        st.toast(f"Synchronized: Row {target_row} | Cover Formed: {new_matter_no}", icon="🔹")
                        
                    st.session_state["pdf_data_store"] = temp_store
                    st.balloons()
            
    with col2:
        if st.button("❌ NO - Abort & Flush Upload Bay", use_container_width=True):
            st.session_state["uploader_key"] += 1
            st.session_state["previous_files"] = []
            st.session_state["pdf_data_store"] = {}
            st.rerun()

# --- 5. INTERACTIVE PRINT BRIDGE ---
if st.session_state["pdf_data_store"]:
    st.markdown("---")
    st.success("🎉 **Data routing completely finalized. Open your physical file templates below to print:**")
    
    for filename, (matter_no, pdf_bytes) in st.session_state["pdf_data_store"].items():
        # Encode binary data to display natively inside browser views if supported
        base64_pdf = base64.b64encode(pdf_bytes).decode('utf-8')
        pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="500" type="application/pdf"></iframe>'
        
        with st.expander(f"📄 View & Print Cover Sheet (Matter No: {matter_no})", expanded=True):
            st.components.v1.html(pdf_display, height=520)
            
            st.download_button(
                label=f"🖨️ Open in Print App / Download Cover Sheet ({matter_no})",
                data=pdf_bytes,
                file_name=f"21Chambers_Cover_{matter_no}.pdf",
                mime="application/pdf",
                use_container_width=True
            )
