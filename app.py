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
# Updated to match the exact file name in the main repository
if os.path.exists("Company Logo.png"):
    st.image("Company Logo.png", width=250)

st.title("📂 21 Chambers Automator")
st.markdown("##### *Streamlined Batch File Processing for Operations*")
st.write("Drag and drop up to **5 open file sheets** simultaneously to log them into the master Google Sheet.")
st.markdown("---")

# --- 1. GOOGLE SHEETS SETUP ---
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
        # Try to access the current month's tab natively
        sheet = workbook.worksheet(SHEET_TAB_NAME)
    except gspread.exceptions.WorksheetNotFound:
        # --- THE ULTIMATE VISUAL CLONER ---
        # Fetch the existing master template tab
        template_sheet = workbook.worksheet("Template")
        
        # Duplicate the template tab exactly into the workbook
        # Copies all dropdowns, colors, borders, and column sizes!
        duplicated_sheet = workbook.duplicate_sheet(
            source_sheet_id=template_sheet.id,
            new_sheet_name=SHEET_TAB_NAME,
            insert_sheet_index=0 # Puts it on the most left slot instantly
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
    
    # Match Type of Work
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
    st.image("Company Logo.png", use_container_width=True)  # Swap with your actual image file name if different!
    st.markdown("### 🛠️ Operation Logistics")
    st.info("This system automates docx intake pipelines directly into the firm master sheet matrix.")
    
    st.markdown("---")
    
    # --- AUTOMATION CREDENTIAL DISPLAY FOR FUTURE OFFICE ADMINS ---
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
            
        st.code(sys_email, language="text")  # Clear, one-click copy box!
    except Exception:
        st.error("Could not resolve system email details.")

st.markdown("---")

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
    unsafe_allow_html=True
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
        st.session_state["uploader_key"] += 1
        st.rerun()

if uploaded_files:
    if len(uploaded_files) > 5:
        st.error("⚠️ System safety cap exceeded. Please upload a maximum of 5 files at a time to prevent server drops.")
    else:
        try:
            with st.spinner(f"Processing batch of {len(uploaded_files)} files..."):
                sheet = get_google_sheet()
                
                # --- 1. FIND THE NEXT AVAILABLE ROW SLOT (FIXED!) ---
                client_values = sheet.col_values(5) 
                target_row = 2
                while target_row <= len(client_values) and client_values[target_row - 1].strip() != "":
                    target_row += 1
                
                # --- 2. INTELLIGENT MONTH-ROLLOVER NUMBER ENGINE ---
                try:
                    matter_nos = sheet.col_values(3)[1:]  
                    valid_numbers = [int(val.strip()) for val in matter_nos if str(val).strip().isdigit()]
                except Exception:
                    valid_numbers = []
                
                if not valid_numbers:
                    try:
                        from datetime import timedelta
                        first_of_this_month = datetime.now().replace(day=1)
                        prev_month_date = first_of_this_month - timedelta(days=15)
                        PREV_SHEET_TAB_NAME = prev_month_date.strftime("%B %Y")
                        
                        # --- FIX: RE-BUILD CREDS PROPERLY FOR LOOK-BACK ---
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
                    except Exception as e:
                        # If lookback genuinely fails, fall back to a safer default or log it
                        valid_numbers = [20260728] 
                
                current_max_matter = max(valid_numbers) if valid_numbers else 20260728
                
                # --- 3. PROCESS EACH FILE IN THE BATCH ---
                for doc_file in uploaded_files:
                    next_index = target_row - 1
                    today_date = datetime.now().strftime("%d %B %Y").lstrip("0")
                    
                    current_max_matter += 1
                    new_matter_no = str(current_max_matter)
                    
                    matter_type, clients, contacts, referral = extract_matter_data(doc_file)
                    
                    # The value structure matching Columns A through I
                    new_row = [
                        next_index,          # Column A: Column 1
                        today_date,          # Column B: Date Opened
                        new_matter_no,       # Column C: Matter No
                        matter_type,         # Column D: Type of Work
                        clients,             # Column E: Client(s)
                        contacts,            # Column F: Contact(s)
                        referral,            # Column G: Referral
                        "Yes",               # Column H: Bill Paid (Yes/No)
                        ""                   # Column I: Closed Date (Kept blank initially)
                    ]
                    
                    # Expanded target range up to column I
                    cell_range = f"A{target_row}:I{target_row}"
                    sheet.update(range_name=cell_range, values=[new_row])
                    
                    st.success(f"✅ Loaded: {doc_file.name} ➡️ Row {target_row} (Matter No: {new_matter_no})")
                    
                    target_row += 1
                    
            st.balloons()

        except Exception as e:
            st.error(f"Error executing automation batch processing: {e}")
