import streamlit as st
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import re
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
import os
import base64
import json
import io

# Enforce clean full-width canvas workspace view
st.set_page_config(page_title="21 Chambers Client List", layout="wide")

if os.path.exists("Company Logo.png"):
    st.image("Company Logo.png", width=250)

st.title("📂 21 Chambers Automator")
st.markdown("##### *Dual-Stream Intake Pipeline & Native Word Document Center*")
st.write("Drag and drop your document below to synchronize the master matrix and prepare perfect physical templates.")
st.markdown("---")

# --- 1. GOOGLE SHEETS CORE CONNECTION ---
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

# --- 2. THE XML ELEMENT INJECTORS FOR CELL PADDING ---
def create_element(name):
    return OxmlElement(name)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = create_element('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = create_element(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

# --- 3. THE HIGH-FIDELITY COVER GENERATOR ENGINE ---
def generate_perfect_docx(matter_no, clients_text, contacts_text, matter_type, date_opened):
    doc = docx.Document()
    
    # Page Setup: A4 parameters mapped to your exact specifications
    for section in doc.sections:
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.295)
        section.right_margin = Inches(0.295)

    # 🏢 TOP TABLE: Spacer & Party/Firm Details
    top_table = doc.add_table(rows=1, cols=2)
    top_table.autofit = False
    top_table.columns[0].width = Inches(1.333)
    top_table.columns[1].width = Inches(6.346)
    
    row_1 = top_table.rows[0]
    row_1.height = Inches(1.567)
    
    row_1.cells[0].width = Inches(1.333)
    cell_b1 = row_1.cells[1]
    cell_b1.width = Inches(6.346)
    set_cell_margins(cell_b1, top=0, bottom=0, left=100, right=0)
    
    p_parties = cell_b1.paragraphs[0]
    p_parties.paragraph_format.line_spacing = 1.15
    p_parties.paragraph_format.space_after = Pt(0)
    
    # Build text flows matching your extraction rules
    lines_list = [l.strip() for l in clients_text.split('\n') if l.strip()]
    app_str = lines_list[0] if len(lines_list) > 0 else "APPLICANT - NIL"
    res_str = lines_list[1] if len(lines_list) > 1 else "RESPONDENT - NIL"
    
    run_app = p_parties.add_run(f"{app_str}\n")
    run_app.font.name = 'Times New Roman'
    run_app.font.size = Pt(20)
    
    # Process and clean contacts field list for presentation display
    contact_lines = [c.strip() for c in contacts_text.split('\n') if c.strip()]
    for c_line in contact_lines:
        run_c = p_parties.add_run(f"{c_line}\n")
        run_c.font.name = 'Times New Roman'
        run_c.font.size = Pt(20)
        
    p_parties.add_run("\n") # Target empty paragraph spacer gap
    
    run_res = p_parties.add_run(f"{res_str}\n")
    run_res.font.name = 'Times New Roman'
    run_res.font.size = Pt(20)
    
    # Append the contacts directly mapped underneath the respondent profile block
    p_firm = cell_b1.add_paragraph()
    p_firm.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_firm.paragraph_format.line_spacing = 1.0
    p_firm.paragraph_format.space_before = Pt(24)
    p_firm.paragraph_format.space_after = Pt(24)
    
    run_firm_name = p_firm.add_run("21 CHAMBERS LLC\n")
    run_firm_name.font.name = 'Times New Roman'
    run_firm_name.font.size = Pt(20)
    run_firm_name.bold = True
    
    run_firm_addr = p_firm.add_run(
        "2 HAVELOCK ROAD #06-17\n"
        "HAVELOCK 2\n"
        "SINGAPORE 059763\n"
        "TEL: 6224 1848     FAX: 6223 3092"
    )
    run_firm_addr.font.name = 'Times New Roman'
    run_firm_addr.font.size = Pt(20)

    # 📉 BOTTOM TABLE: Matter Specifications Matrix
    bottom_table = doc.add_table(rows=4, cols=2)
    bottom_table.autofit = False
    bottom_table.columns[0].width = Inches(1.596)
    bottom_table.columns[1].width = Inches(6.083)
    
    row_heights = [Inches(0.792), Inches(1.133), Inches(0.208), Inches(0.208)]
    
    # Map label options dynamically matching the exact parsed case metrics
    full_matter_name = "Uncontested Divorce"
    if matter_type == "CD": full_matter_name = "Contested Divorce"
    elif matter_type == "Annulment": full_matter_name = "Annulment"
    elif matter_type == "Variation": full_matter_name = "Variation"
    elif matter_type == "Others": full_matter_name = "Others"

    row_data = [
        {"label": "SUBJECT\nMATTER", "value": full_matter_name, "bold_label": True},
        {"label": "FILE", "value": f"{matter_no}\nOpening date: {date_opened}\nClosure date:", "bold_label": False},
        {"label": "Legal Fee", "value": "CASH", "bold_label": False},
        {"label": "Remarks", "value": "", "bold_label": False}
    ]
    
    for idx, row in enumerate(bottom_table.rows):
        row.height = row_heights[idx]
        
        cell_a = row.cells[0]
        cell_a.width = Inches(1.596)
        set_cell_margins(cell_a, top=60, bottom=60, left=100, right=60)
        p_a = cell_a.paragraphs[0]
        p_a.paragraph_format.space_after = Pt(0)
        run_a = p_a.add_run(row_data[idx]["label"])
        run_a.font.name = 'Times New Roman'
        run_a.font.size = Pt(20)
        if row_data[idx]["bold_label"]:
            run_a.bold = True
            
        cell_b = row.cells[1]
        cell_b.width = Inches(6.083)
        set_cell_margins(cell_b, top=60, bottom=60, left=100, right=60)
        p_b = cell_b.paragraphs[0]
        p_b.paragraph_format.space_after = Pt(0)
        p_b.paragraph_format.line_spacing = 1.15
        run_b = p_b.add_run(row_data[idx]["value"])
        run_b.font.name = 'Times New Roman'
        run_b.font.size = Pt(20)

    # 🤯 THE GIANT FOOTER: Arial Bold 109 Signpost Block
    p_mega = doc.add_paragraph()
    p_mega.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_mega.paragraph_format.space_before = Pt(54)
    p_mega.paragraph_format.space_after = Pt(0)
    
    run_mega = p_mega.add_run(matter_no)
    run_mega.font.name = 'Arial'
    run_mega.font.size = Pt(109)
    run_mega.bold = True
    
    # Save directly down to a binary file stream buffer
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream.getvalue()

# --- 4. RAW INPUT DOCX SCRAPING ENGINE ---
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

# --- 5. WORKSPACE FLOW SYSTEMS ---
if "uploader_key" not in st.session_state: st.session_state["uploader_key"] = 0
if "previous_files" not in st.session_state: st.session_state["previous_files"] = []
if "docx_binary_store" not in st.session_state: st.session_state["docx_binary_store"] = {}

uploaded_files = st.file_uploader(
    "Drag and drop Open File Sheets (.docx) here", 
    type=["docx"], accept_multiple_files=True, key=f"uploader_{st.session_state['uploader_key']}"
)

current_file_names = [f.name for f in uploaded_files] if uploaded_files else []
if current_file_names != st.session_state["previous_files"]:
    st.session_state["previous_files"] = current_file_names
    st.session_state["docx_binary_store"] = {}

if uploaded_files:
    st.markdown("---")
    st.subheader("🖨️ Production Queue Confirmation")
    st.write("Documents read successfully. Sync data matrix rows and prepare physical office templates?")
    
    c1, c2 = st.columns(2)
    with c1:
        if st.button("✅ YES - Execute Streams & Generate Layouts", use_container_width=True, type="primary"):
            if not st.session_state["docx_binary_store"]:
                with st.spinner("⚡ Running full-scale dual execution arrays..."):
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
                        t_date = datetime.now().strftime("%d %B %Y").lstrip("0")
                        
                        m_type, cls, cnt, ref = extract_matter_data(f)
                        
                        # Stream 1: Update Google Sheets row data
                        new_row = [next_idx, t_date, new_no, m_type, cls, cnt, ref, "Yes", ""]
                        sheet.update(range_name=f"A{target_row}:I{target_row}", values=[new_row])
                        
                        # Stream 2: Form perfect template layout using your exact parameters
                        docx_bytes = generate_perfect_docx(new_no, cls, cnt, m_type, t_date)
                        temp_store[f.name] = (new_no, docx_bytes)
                        st.toast(f"Synchronized Case File Matrix: Matter {new_no}", icon="🔹")
                        
                    st.session_state["docx_binary_store"] = temp_store
                    st.balloons()
                    
    with c2:
        if st.button("❌ NO - Abort Operational Batch", use_container_width=True):
            st.session_state["uploader_key"] += 1
            st.session_state["previous_files"] = []
            st.session_state["docx_binary_store"] = {}
            st.rerun()

# --- 6. CLEAN AUTOMATED DOWNLOAD MATRIX ---
if st.session_state["docx_binary_store"]:
    st.markdown("---")
    st.success("🎉 **Data routing completely finalized. Open your physical file templates below:**")
    
    for fname, (m_no, docx_bytes) in st.session_state["docx_binary_store"].items():
        st.download_button(
            label=f"📝 Save & Open Perfect Template File Cover (Matter File Number: {m_no})",
            data=docx_bytes,
            file_name=f"21Chambers_Cover_{m_no}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
