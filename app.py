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

# ReportLab positioning elements
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# Force clear full-width canvas layout workspace view
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

# --- 2. TYPOGRAPHY POSITIONING ENGINE ---
def generate_perfect_pdf(matter_no, clients_text, contacts_text, matter_type, date_opened):
    buffer = io.BytesIO()
    
    # Page Setup: A4 with your exact 0.295" margins (21.24 points)
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        leftMargin=21.24, rightMargin=21.24,
        topMargin=36.0, bottomMargin=36.0
    )
    
    printable_width = A4[0] - 42.48 
    
    # FIX: We set explicit 'leading' matching the massive 20pt font so text lines never collide
    style_normal_20 = ParagraphStyle('Norm20', fontName='Times-Roman', fontSize=20, leading=26, alignment=TA_LEFT)
    style_bold_20 = ParagraphStyle('Bold20', fontName='Times-Bold', fontSize=20, leading=26, alignment=TA_LEFT)
    style_firm_title = ParagraphStyle('FirmTitle', fontName='Times-Bold', fontSize=20, leading=26, alignment=TA_CENTER)
    style_firm_body = ParagraphStyle('FirmBody', fontName='Times-Roman', fontSize=20, leading=26, alignment=TA_CENTER)
    
    # The Mega Footer Code (Arial Bold 109pt)
    style_giant_foot = ParagraphStyle('GiantFoot', fontName='Helvetica-Bold', fontSize=109, leading=115, alignment=TA_CENTER)
    
    story = []
    
    # ----------------------------------------------------
    # 🏢 TOP SECTION: Spacing & Party / Firm Layout
    # ----------------------------------------------------
    client_lines = [line.strip() for line in clients_text.split('\n') if line.strip()]
    contact_lines = [line.strip() for line in contacts_text.split('\n') if line.strip()]
    
    right_cell_elements = []
    
    # Process Applicant Block
    if len(client_lines) > 0:
        right_cell_elements.append(Paragraph(client_lines[0], style_normal_20))
    if len(contact_lines) > 0:
        right_cell_elements.append(Paragraph(contact_lines[0], style_normal_20))
        
    right_cell_elements.append(Spacer(1, 14)) # Clean vertical separation gap
    
    # Process Respondent Block
    if len(client_lines) > 1:
        right_cell_elements.append(Paragraph(client_lines[1], style_normal_20))
    if len(contact_lines) > 1:
        right_cell_elements.append(Paragraph(contact_lines[1], style_normal_20))
        
    # Firm Letterhead Section
    right_cell_elements.append(Spacer(1, 24))
    right_cell_elements.append(Paragraph("21 CHAMBERS LLC", style_firm_title))
    right_cell_elements.append(Spacer(1, 4))
    right_cell_elements.append(Paragraph("2 HAVELOCK ROAD #06-17<br/>HAVELOCK 2<br/>SINGAPORE 059763", style_firm_body))
    right_cell_elements.append(Spacer(1, 4))
    right_cell_elements.append(Paragraph("TEL: 6224 1848 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; FAX: 6223 3092", style_firm_body))
    
    # Left Spacer (1.333") vs Right Data Info Column (6.346") converted to layout units
    col_w_left = 1.333 * 72
    col_w_right = 6.346 * 72
    
    # FIX: Removed the rigid height constraint on row Heights! The table now auto-expands cleanly
    top_table = Table([["", right_cell_elements]], colWidths=[col_w_left, col_w_right])
    top_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('LEFTPADDING', (1,0), (1,0), 7.2), # 100 dxa left cell padding indentation
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 0),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
    ]))
    story.append(top_table)
    story.append(Spacer(1, 35))
    
    # ----------------------------------------------------
    # 📉 MIDDLE SECTION: Matter Specifications Matrix
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
    
    # Label col width (1.596") vs value col width (6.083")
    b_col1 = 1.596 * 72
    b_col2 = 6.083 * 72
    
    # FIX: No forced cell height limitations, allowing perfect dynamic placement
    bottom_table = Table(matrix_rows, colWidths=[b_col1, b_col2])
    bottom_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING', (0,0), (-1,-1), 7.2), # Exact 100 dxa spacing
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(bottom_table)
    story.append(Spacer(1, 54)) 
    
    # ----------------------------------------------------
    # 🤯 BOTTOM SECTION: The Giant Footer Display
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

# --- 4. STREAMLIT FRAMEWORK FLOWS ---
if "uploader_key" not in st.session_state: st.session_state["uploader_key"] = 0
if "previous_files" not in st.session_state: st.session_state["previous_files"] = []
if "pdf_binary_store" not in st.session_state: st.session_state["pdf_binary_store"] = {}

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
                with st.spinner("⚡ Running layout updates..."):
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
                        
                        # Generate the newly aligned text document layout
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
    st.success("🎉 **Data routing finalized. Download the adjusted text layout below:**")
    for fname, (m_no, pdf_bytes) in st.session_state["pdf_binary_store"].items():
        st.download_button(
            label=f"🖨️ Download Alignment Check PDF (Matter: {m_no})",
            data=pdf_bytes,
            file_name=f"21Chambers_TextCheck_{m_no}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
