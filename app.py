import streamlit as st
import docx
import re
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
import os
import base64
import json

st.set_page_config(page_title="21 Chambers Client List", layout="centered")

# --- FORCE SYSTEM RE-RENDER INSULATION ---
st.markdown("", unsafe_allow_html=True)

# --- UI INTERFACE GRAPHICS & LOGO ---
if os.path.exists("Company Logo.png"):
    st.image("Company Logo.png", width=250)

st.title("📂 21 Chambers Automator")
st.markdown("##### *Dual-Stream Intake Pipeline & Native Print Center*")
st.write("Drag and drop your document below to synchronize the master matrix and prepare print layouts.")
st.markdown("---")

# --- 1. GOOGLE SHEETS SETUP ---
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

# --- 2. HIGH-FIDELITY HTML/CSS COVER GENERATOR ---
def generate_print_html(matter_no, clients_text, contacts_text, matter_type, date_opened):
    client_lines = [line.strip() for line in clients_text.split('\n') if line.strip()]
    contact_lines = [line.strip() for line in contacts_text.split('\n') if line.strip()]
    
    top_box_content = ""
    for line in client_lines:
        if "APPLICANT" in line or "RESPONDENT" in line:
            top_box_content += f"<div style='font-weight: bold;'>{line}</div>"
        else:
            top_box_content += f"<div>{line}</div>"
    for line in contact_lines:
        top_box_content += f"<div>{line}</div>"
        
    subject_label = "GLA for Estate of ..."
    if client_lines:
        clean_name = client_lines[0].replace("APPLICANT - ", "")
        subject_label = f"{matter_type} for {clean_name}"

    html_layout = f"""
    <html>
    <head>
    <style>
        body {{
            font-family: 'Arial', 'Calibri', sans-serif;
            color: #000000;
            line-height: 1.25;
            font-size: 13pt;
            background-color: #ffffff;
            padding: 20px;
        }}
        .top-box {{
            border: 1.5px solid #000000;
            padding: 14px;
            margin-bottom: 28pt;
        }}
        .center-text {{
            text-align: center;
        }}
        .firm-header {{
            margin-bottom: 36pt;
        }}
        .firm-title {{
            font-size: 20pt;
            font-weight: bold;
            margin-bottom: 4pt;
            letter-spacing: 0.5px;
        }}
        .firm-body {{
            font-size: 13pt;
            margin-bottom: 4pt;
        }}
        table.matrix-table {{
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 48pt;
        }}
        table.matrix-table td {{
            border: 1.5px solid #000000;
            padding: 10px;
            vertical-align: top;
            font-size: 13pt;
            color: #000000;
        }}
        table.matrix-table td.label {{
            font-weight: bold;
            width: 25%;
        }}
        .giant-footer {{
            font-size: 76pt;
            font-weight: bold;
            text-align: center;
            margin-top: 40pt;
            letter-spacing: 1px;
        }}
    </style>
    </head>
    <body>
        <div class="top-box">{top_box_content}</div>
        <div class="center-text firm-header">
            <div class="firm-title">21 CHAMBERS LLC</div>
            <div class="firm-body">2 HAVELOCK ROAD #06-17<br/>HAVELOCK 2<br/>SINGAPORE 059763</div>
            <div class="firm-body">TEL: 6224 1848 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; FAX: 6223 3092</div>
        </div>
        <table class="matrix-table">
            <tr><td class="label">SUBJECT MATTER</td><td>{subject_label}</td></tr>
            <tr><td class="label">FILE</td><td><strong>{matter_no}</strong><br/>Opening date: {date_opened}<br/>Closure date:</td></tr>
            <tr><td class="label">Legal Fee</td><td>CASH</td></tr>
            <tr><td class="label">Remarks</td><td></td></tr>
        </table>
        <div class="giant-footer">{matter_no}</div>
    </body>
    </html>
    """
    return html_layout

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
    st.info("Dual-Stream Processing Node Active")

# --- 4. RUNTIME SYSTEM STATES ---
if "uploader_key" not in st.session_state: st.session_state["uploader_key"] = 0
if "previous_files" not in st.session_state: st.session_state["previous_files"] = []
if "processed_html_store" not in st.session_state: st.session_state["processed_html_store"] = {}

uploaded_files = st.file_uploader(
    "Upload Intake Sheets (.docx)", 
    type=["docx"], accept_multiple_files=True, key=f"uploader_{st.session_state['uploader_key']}"
)

current_file_names = [f.name for f in uploaded_files] if uploaded_files else []
if current_file_names != st.session_state["previous_files"]:
    st.session_state["previous_files"] = current_file_names
    st.session_state["processed_html_store"] = {}

# --- 5. THE FAIL-SAFE QUEUE SYSTEM ---
if uploaded_files:
    st.markdown("---")
    st.subheader("🖨️ Form Pipeline Confirmation")
    
    c1, c2 = st.columns(2)
    with c1:
        if st.button("✅ RUN AUTO-LOG & VIEW COVERS", use_container_width=True, type="primary"):
            if not st.session_state["processed_html_store"]:
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
                    new_row = [next_idx, t_date, new_no, m_type, cls, cnt, ref, "Yes", ""]
                    sheet.update(range_name=f"A{target_row}:I{target_row}", values=[new_row])
                    
                    html_view = generate_print_html(new_no, cls, cnt, m_type, t_date)
                    temp_store[f.name] = (new_no, html_view)
                    
                st.session_state["processed_html_store"] = temp_store
                st.balloons()
                
    with c2:
        if st.button("❌ CANCEL & WIPE BAY", use_container_width=True):
            st.session_state["uploader_key"] += 1
            st.session_state["previous_files"] = []
            st.session_state["processed_html_store"] = {}
            st.rerun()

# --- 6. NATIVE WEB PRINT DISPATCHERS ---
if st.session_state["processed_html_store"]:
    st.markdown("---")
    for fname, (m_no, html_code) in st.session_state["processed_html_store"].items():
        with st.expander(f"🖨️ COVER SHEET PRINT MODULE - MATTER {m_no}", expanded=True):
            st.components.v1.html(html_code, height=650, scrolling=True)
            
            b64 = base64.b64encode(html_code.encode('utf-8')).decode('utf-8')
            print_button_script = f"""
                <a href="data:text/html;base64,{b64}" target="_blank" style="text-decoration:none;">
                    <button style="width:100%; background-color:#FF4B4B; color:white; border:none; padding:14px; font-size:16px; font-weight:bold; border-radius:8px; cursor:pointer;">
                        🖨️ OPEN PRINT TAB FOR MATTER {m_no}
                    </button>
                </a>
            """
            st.components.v1.html(print_button_script, height=60)
