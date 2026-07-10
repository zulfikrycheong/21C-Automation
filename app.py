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

# --- UI INTERFACE GRAPHICS & LOGO ---
if os.path.exists("Company Logo.png"):
    st.image("Company Logo.png", width=250)

st.title("📂 21 Chambers Automator")
st.markdown("##### *Streamlined Batch File Processing for Operations*")
st.write("Drag and drop up to **5 open file sheets** simultaneously to log them into the master Google Sheet.")
st.markdown("---")

# --- 1. GOOGLE SHEETS SETUP ---
# Kept strictly as your original master name so it doesn't lose connection to your file!
GOOGLE_SHEET_NAME = "Lazy Automation"  
SHEET_TAB_NAME = "August 2026"

def get_google_sheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    
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
        # --- THE ULTIMATE VISUAL CLONER ---
        template_sheet = workbook.worksheet("Template")
        
        duplicated_sheet = workbook.duplicate_sheet(
            source_sheet_id=template_sheet.id,
            new_sheet_name=SHEET_TAB_NAME,
            insert_sheet_index=0 
        )
        sheet = duplicated_sheet
        
    return sheet
    
# --- 2. DOCX PARSING LOGIC ---
def extract_matter_data(doc_path):
    doc = docx.Document(doc_path)
    
    text_lines = []
    for para in doc.paragraphs:
        if para.text.strip():
            text_lines.append(para.text)
            
    for table in doc.tables:
        for row in table.rows:
            row_text = [cell.text.strip() for cell in row.cells if cell.text.strip()]
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
        if ref_match:
            extracted_ref = ref_match.group(1).lower()
            if "jav" in extracted_ref:
                referral = "Javern"
        elif "jav" in full_text.lower():
            referral = "Javern"
            
    return matter_type, clients_field, contacts_field, referral

# --- SIDEBAR DEPLOYMENT ARCHITECTURE ---
with st.sidebar:
    st.image("Company Logo.png", use_container_width=True)  
    st.markdown("### 🛠️ Operation Logistics")
    st.info("This system automates docx intake pipelines directly into the firm master sheet matrix.")
    
    st.markdown("---")
    
    st.markdown("### 📋 System Access Email")
    st.caption("If the firm creates a brand new master Google Sheet file, you MUST share the new sheet with this email as an **Editor** or the app will fail:")
    
    try:
        if os.path.exists("credentials.json"):
            with open("credentials.json", "r") as f:
                c_data = json.load(f)
                sys_email = c_data.get("client_email", "Not found")
        else:
            encoded_str = st.secrets["encoded_creds"]
            decoded_bytes = base64.b64decode(encoded_str)
            creds_dict = json.loads(decoded_bytes)
            sys_email = creds_dict.get("client_email", "Not found")
            
        st.code(sys_email, language="text")  
    except Exception:
        st.error("Could not resolve system email details.")

st.markdown("---")

# --- 3. RUNTIME BATCH LOGIC ---

st.markdown(
    """
    <style>
    [data-testid="stFileUploader"] {
        padding: 20px 0px;
    }
    [data-testid="stFileUploaderDropzone"] {
        padding: 50px 20px;
        min-height: 220px;
    }
    </style>
    """,
    unsafe_allow_html=True
)

if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0
if "previous_files" not in st.session_state:
    st.session_state["previous_files"] = []

uploaded_files = st.file_uploader(
    "Drag and drop Open File Sheets (.docx) here", 
    type=["docx"], 
    accept_multiple_files=True,
    key=f"uploader_{st.session_state['uploader_key']}"
)

# --- CLEAN AUTOMATED FLUSH MECHANISM ---
current_file_names = [f.name for f in uploaded_files] if uploaded_files else []

# Instantly clears out outdated UI logs from the screen if the batch list shifts
if current_file_names != st.session_state["previous_files"]:
    st.session_state["previous_files"] = current_file_names
    # Allows natural processing layout initialization without lifecycle interruptions

if uploaded_files:
    if st.button("🧹 Clear Upload Bay", use_container_width=True):
        st.session_state["uploader_key"] += 1
        st.session_state["previous_files"] = []
        st.rerun()

if uploaded_files:
    if len(uploaded_files) > 5:
        st.error("⚠️ System safety cap exceeded. Please upload a maximum of 5 files at a time to prevent server drops.")
    else:
        try:
            with st.spinner(f"Processing batch of {len(uploaded_files)} files..."):
                sheet = get_google_sheet()
                
                # --- 2. INTELLIGENT MONTH/YEAR ROLLOVER SEQUENCING ---
                try:
                    matter_nos = sheet.col_values(3)[1:]  
                    valid_numbers = [int(val.strip()) for val in matter_nos if str(val).strip().isdigit()]
                except Exception:
                    valid_numbers = []
                
                if not valid_numbers:
                    if datetime.now().month == 1:
                        current_year_str = datetime.now().strftime("%Y")
                        current_max_matter = int(f"{current_year_str}0000")
                    else:
                        try:
                            from datetime import timedelta
                            first_of_this_month = datetime.now().replace(day=1)
                            prev_month_date = first_of_this_month - timedelta(days=15)
                            PREV_SHEET_TAB_NAME = prev_month_date.strftime("%B %Y")
                            
                            scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
                            if os.path.exists("credentials.json"):
                                lookback_creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
                            else:
                                encoded_str = st.secrets["encoded_creds"]
                                decoded_bytes = base64.b64decode(encoded_str)
                                creds_dict = json.loads(decoded_bytes)
                                lookback_creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
                                
                            client = gspread.authorize(lookback_creds)
                            prev_sheet = client.open(GOOGLE_SHEET_NAME).worksheet(PREV_SHEET_TAB_NAME)
                            prev_matter_nos = prev_sheet.col_values(3)[1:]
                            valid_numbers = [int(val.strip()) for val in prev_matter_nos if str(val).strip().isdigit()]
                        except Exception:
                            valid_numbers = [20260728] 
                
                current_max_matter = max(valid_numbers) if valid_numbers else 20260728
                
                # --- 3. PROCESS EACH FILE IN THE BATCH ---
                for doc_file in uploaded_files:
                    matter_nos = sheet.col_values(3) 
                    max_num = -1
                    max_row_index = 1 
                    
                    for idx, val in enumerate(matter_nos):
                        clean_val = str(val).strip()
                        if clean_val.isdigit():
                            num = int(clean_val)
                            if num > max_num:
                                max_num = num
                                max_row_index = idx + 1 
                    
                    if max_num == -1:
                        target_row = 2
                        if datetime.now().month == 1:
                            current_year_str = datetime.now().strftime("%Y")
                            current_max_matter = int(f"{current_year_str}0000")
                        else:
                            current_max_matter = 20260728 
                    else:
                        target_row = max_row_index + 1
                        current_max_matter = max_num
                    
                    current_max_matter += 1
                    new_matter_no = str(current_max_matter)
                    next_index = target_row - 1
                    today_date = datetime.now().strftime("%d %B %Y").lstrip("0")
                    
                    matter_type, clients, contacts, referral = extract_matter_data(doc_file)
                    
                    new_row = [
                        next_index, today_date, new_matter_no, matter_type, 
                        clients, contacts, referral, "Yes", ""
                    ]
                    
                    cell_range = f"A{target_row}:I{target_row}"
                    sheet.update(range_name=cell_range, values=[new_row])
                    
                    st.success(f"✅ Loaded: {doc_file.name} ➡️ Row {target_row} (Matter No: {new_matter_no})")
                    
            st.balloons()

        # --- THE AUTOMATIC RESET TRICK ---
            # Automatically increments the uploader widget key so it clears out the files for the next batch run!
            st.session_state["uploader_key"] += 1
            st.session_state["previous_files"] = []
            st.rerun()
        
        except Exception as e:
            st.error(f"Error executing automation batch processing: {e}")
