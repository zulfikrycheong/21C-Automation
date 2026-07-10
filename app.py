import streamlit as st
import docx
import re
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
import os
import base64
import json

st.set_page_config(page_title="Firm Master Sheet Automator", layout="centered")

# --- UI INTERFACE GRAPHICS & LOGO ---
# Updated to match your exact file name in the main repository
if os.path.exists("Company Logo.png"):
    st.image("Company Logo.png", width=250)

st.title("📂 Law Firm Intake Automator")
st.markdown("##### *Streamlined Batch File Processing for Operations*")
st.write("Drag and drop up to **5 open file sheets** simultaneously to log them into the master Google Sheet.")
st.markdown("---")

# --- 1. GOOGLE SHEETS SETUP ---
GOOGLE_SHEET_NAME = "Lazy Automation"  
SHEET_TAB_NAME = datetime.now().strftime("%B %Y")

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
    sheet = client.open(GOOGLE_SHEET_NAME).worksheet(SHEET_TAB_NAME)
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

# --- 3. RUNTIME BATCH LOGIC ---

# Inject custom CSS to enlarge the drag-and-drop upload bay surface area
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
    unsafe_html=True
)

# Initialize a session state key for the file uploader to allow manual clearing
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0

# The expanded file uploader bay
uploaded_files = st.file_uploader(
    "Drag and drop Open File Sheets (.docx) here", 
    type=["docx"], 
    accept_multiple_files=True,
    key=f"uploader_{st.session_state['uploader_key']}"
)

# Add a clean clear button right below the upload bay if files are present
if uploaded_files:
    if st.button("🧹 Clear Upload Bay", use_container_width=True):
        # Incrementing the key forces Streamlit to completely destroy the old widget and draw a fresh, empty one
        st.session_state["uploader_key"] += 1
        st.rerun()

if uploaded_files:
    # Strict safeguard check to enforce the 5-file cap
    if len(uploaded_files) > 5:
        st.error("⚠️ System safety cap exceeded. Please upload a maximum of 5 files at a time to prevent server drops.")
    else:
        try:
            with st.spinner(f"Processing batch of {len(uploaded_files)} files..."):
                sheet = get_google_sheet()
                
                # Load sheet structures outside the loop to minimize API overhead
                matter_nos = sheet.col_values(3)[1:]  
                valid_numbers = [int(val.strip()) for val in matter_nos if str(val).strip().isdigit()]
                current_max_matter = max(valid_numbers) if valid_numbers else 20260622
                
                client_values = sheet.col_values(5) 
                target_row = 2
                while target_row <= len(client_values) and client_values[target_row - 1].strip() != "":
                    target_row += 1
                
                # Process each file one by one in order
                for doc_file in uploaded_files:
                    next_index = target_row - 1
                    today_date = datetime.now().strftime("%d %B %Y").lstrip("0")
                    
                    # Increment matter number natively per file loop iteration
                    current_max_matter += 1
                    new_matter_no = str(current_max_matter)
                    
                    matter_type, clients, contacts, referral = extract_matter_data(doc_file)
                    
                    new_row = [
                        next_index, today_date, new_matter_no, matter_type, 
                        clients, contacts, referral, "Yes"
                    ]
                    
                    cell_range = f"A{target_row}:H{target_row}"
                    sheet.update(range_name=cell_range, values=[new_row])
                    
                    # Visually render success tracking elements for each logged document
                    st.success(f"✅ Loaded: {doc_file.name} ➡️ Row {target_row} (Matter No: {new_matter_no})")
                    
                    # Push target row tracker down by 1 for the next file in line
                    target_row += 1
                    
            st.balloons() # Batch complete celebration!

        except Exception as e:
            st.error(f"Error executing automation batch processing: {e}")
