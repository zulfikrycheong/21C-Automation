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

# We import the lightweight fpdf2 library directly inside the workspace
try:
    from fpdf import FPDF
except ModuleNotFoundError:
    import subprocess
    import sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "fpdf2"])
    from fpdf import FPDF

# FORCE STREAMLIT TO RENDER IN FULL-WIDTH FLAT LAYOUT
st.set_page_config(page_title="21 Chambers Client List", layout="wide")

# --- UI INTERFACE GRAPHICS & LOGO ---
if os.path.exists("Company Logo.png"):
    st.image("Company Logo.png", width=250)

st.title("📂 21 Chambers Automator")
st.markdown("##### *Dual-Stream Intake Pipeline & Native Vector PDF Center*")
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

# --- 2. VECTOR PDF BLUEPRINT GENERATOR (A4 Pixel-Perfect Replication) ---
class CoverSheetPDF(FPDF):
    def generate_cover(self, matter_no, clients, contacts, matter_type, date_opened):
        self.add_page()
        self.set_margins(20, 20, 20)
        self.set_auto_page_break(False)
        
        # Calculate clean width metrics
        page_w = self.w - 40
        
        # Extract individual name profiles safely
        client_lines = [line.strip() for line in clients.split('\n') if line.strip()]
        contact_lines = [line.strip() for line in contacts.split('\n') if line.strip()]
        
        # A. TOP CONTAINER BOX (Calculate total inner text height dynamically)
        self.set_font("Arial", size=13)
        box_text_lines = client_lines + contact_lines
        box_height = (len(box_text_lines) * 7) + 14
        
        # Draw the bounding rectangle outline
        self.rect(20, 20, page_w, box_height)
        self.set_xy(25, 25)
        
        for line in box_text_lines:
            if "APPLICANT" in line or "RESPONDENT" in line:
                self.set_font("Arial", "B", size=13)
            else:
                self.set_font("Arial", size=13)
            self.cell(page_w - 10, 6, line, ln=True)
            self.set_x(25)
            
        # Move system cursor past the box margin boundaries
        self.set_y(20 + box_height + 25)
        
        # B. CENTRAL FIRM EMBLEM BLOCK
        self.set_font("Arial", "B", size=20)
        self.cell(page_w, 8, "21 CHAMBERS LLC", align="C", ln=True)
        self.ln(2)
        
        self.set_font("Arial", size=13)
        self.cell(page_w, 6, "2 HAVELOCK ROAD #06-17", align="C", ln=True)
        self.cell(page_w, 6, "HAVELOCK 2", align="C", ln=True)
        self.cell(page_w, 6, "SINGAPORE 059763", align="C", ln=True)
        self.ln(2)
        self.cell(page_w, 6, "TEL: 6224 1848       FAX: 6223 3092", align="C", ln=True)
        self.ln(30)
        
        # C. CORE METADATA MATRIX GRID Table Mapping
        clean_name = client_lines[0].replace("APPLICANT - ", "") if client_lines else "NIL"
        subject_label = f"{matter_type} for {clean_name}"
        
        # Establish table widths at a sharp 25% / 75% geometric balance
        col1_w = page_w * 0.25
        col2_w = page_w * 0.75
        
        # Row 1: Subject Matter
        self.set_font("Arial", "B", size=13)
        self.cell(col1_w, 12, "SUBJECT MATTER", border=1)
        self.set_font("Arial", size=13)
        self.cell(col2_w, 12, f" {subject_label}", border=1, ln=True)
        
        # Row 2: File Tracking Metadata
        self.set_font("Arial", "B", size=13)
        # Store structural positioning to accommodate multi-line address blocks natively inside cells
        start_x = self.get_x()
        start_y = self.get_y()
        self.cell(col1_w, 20, "FILE", border=1)
        
        self.set_xy(start_x + col1_w, start_y)
        self.set_font("Arial", "B", size=13)
        self.cell(col2_w, 20, "", border=1) # Outer cell bounding box
        self.set_xy(start_x + col1_w + 2, start_y + 2)
        self.cell(col2_w - 4, 5, matter_no, ln=True)
        self.set_x(start_x + col1_w + 2)
        self.set_font("Arial", size=13)
        self.cell(col2_w - 4, 5, f"Opening date: {date_opened}", ln=True)
        self.set_x(start_x + col1_w + 2)
        self.cell(col2_w - 4, 5, "Closure date:", ln=True)
        
        self.set_xy(start_x, start_y + 20)
        
        # Row 3: Financial Fees
        self.set_font("Arial", "B", size=13)
        self.cell(col1_w, 12, "Legal Fee", border=1)
        self.set_font("Arial", size=13)
        self.cell(col2_w, 12, " CASH", border=1, ln=True)
        
        # Row 4: Remarks Blank Cell
        self.set_font("Arial", "B", size=13)
        self.cell(col1_w, 12, "Remarks", border=1)
        self.set_font("Arial", size=13)
        self.cell(col2_w, 12, "", border=1, ln=True)
        self.ln(35)
        
        # D. GIANT FOOTER CODE TRACKING DISPLAY (Set to high-visibility bold scale)
        self.set_font("Arial", "B", size=76)
        self.cell(page_w, 25, matter_no, align="C", ln=True)
        
        return self.output()

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

# --- SIDEBAR ACCESS CONTROL PANEL ---
with st.sidebar:
    st.image("Company Logo.png", use_container_width=True)  
    st.markdown("### 🛠️ Operation Logistics")
    st.info("Full-Width Dual Stream Core Engine Active.")

# --- 4. RUNTIME SYSTEM CONFIGURATIONS ---
if "uploader_key" not in st.session_state: st.session_state["uploader_key"] = 0
if "previous_files" not in st.session_state: st.session_state["previous_files"] = []
if "pdf_binary_store" not in st.session_state: st.session_state["pdf_binary_store"] = {}

# Forced flat execution canvas area layout
uploaded_files = st.file_uploader(
    "Drag and drop Open File Sheets (.docx) here", 
    type=["docx"], accept_multiple_files=True, key=f"uploader_{st.session_state['uploader_key']}"
)

current_file_names = [f.name for f in uploaded_files] if uploaded_files else []
if current_file_names != st.session_state["previous_files"]:
    st.session_state["previous_files"] = current_file_names
    st.session_state["pdf_binary_store"] = {}

# --- 5. THE FAIL-SAFE PRODUCTION BAY QUEUE ---
if uploaded_files:
    st.markdown("---")
    st.subheader("🖨️ Production Queue Confirmation")
    st.write("Documents read successfully. Sync data matrix rows and prepare physical print layouts?")
    
    c1, c2 = st.columns(2)
    with c1:
        if st.button("✅ YES - Execute Streams & Generate Layouts", use_container_width=True, type="primary"):
            if not st.session_state["pdf_binary_store"]:
                with st.spinner("⚡ Processing dual data arrays..."):
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
                        
                        # Stream 1: Update the remote Google Sheets data cells
                        new_row = [next_idx, t_date, new_no, m_type, cls, cnt, ref, "Yes", ""]
                        sheet.update(range_name=f"A{target_row}:I{target_row}", values=[new_row])
                        
                        # Stream 2: Render crisp vector PDF boundaries directly in server memory cache
                        pdf_engine = CoverSheetPDF()
                        pdf_output_bytes = pdf_engine.generate_cover(new_no, cls, cnt, m_type, t_date)
                        
                        temp_store[f.name] = (new_no, pdf_output_bytes)
                        st.toast(f"Synchronized Case File Matrix: Matter {new_no}", icon="🔹")
                        
                    st.session_state["pdf_binary_store"] = temp_store
                    st.balloons()
                    
    with c2:
        if st.button("❌ NO - Abort Operational Batch", use_container_width=True):
            st.session_state["uploader_key"] += 1
            st.session_state["previous_files"] = []
            st.session_state["pdf_binary_store"] = {}
            st.rerun()

# --- 6. CRISP ARCHITECTURAL DOWNLOAD CONSOLES ---
if st.session_state["pdf_binary_store"]:
    st.markdown("---")
    st.success("🎉 **Data routing completely finalized. Action individual hardware print streams below:**")
    
    for fname, (m_no, pdf_bytes) in st.session_state["pdf_binary_store"].items():
        # High-utility individual print module button row
        st.download_button(
            label=f"🖨️ Open & Print Cover Sheet (Matter File Number: {m_no})",
            data=pdf_bytes,
            file_name=f"21Chambers_Cover_{m_no}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
