import streamlit as st
import docx
import re
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="Firm Master Sheet Automator", layout="centered")
st.title("📂 Law Firm Intake Automator")
st.write("Drag and drop an open file sheet to automatically log it into the master Google Sheet.")

# --- 1. GOOGLE SHEETS SETUP ---
GOOGLE_SHEET_NAME = "Lazy Automation"  
SHEET_TAB_NAME = "July 2026"  

def get_google_sheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
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
    
    # Match Type of Work
    if "uncontested divorce" in full_text.lower():
        matter_type = "UD"
    elif "contested divorce" in full_text.lower():
        matter_type = "CD"
    else:
        matter_type = "Others"
    
    # Extract Names safely
    app_match = re.search(r"Applicant\s*–\s*([^,\n\d]+)", full_text, re.IGNORECASE)
    res_match = re.search(r"Respondent\s*–\s*([^–,\n\d]+)", full_text, re.IGNORECASE)
    
    applicant = "APPLICANT - NIL"
    if app_match:
        name = re.sub(r"\s*-\s*upload.*", "", app_match.group(1), flags=re.IGNORECASE).strip().upper()
        applicant = f"APPLICANT - {name}"
        
    respondent = "RESPONDENT - NIL"
    if res_match:
        name = re.sub(r"\s*-\s*upload.*", "", res_match.group(1), flags=re.IGNORECASE).strip().upper()
        respondent = f"RESPONDENT - {name}"
        
    clients_field = f"{applicant}\n{respondent}"
    
    # --- BULLETPROOF CONTACT EXTRACTION ---
    # Captures any 8-digit Singapore phone number sequence (isolated or after text labels)
    mobiles = re.findall(r"\b[89]\d{7}\b", full_text)
    
    # Captures any valid email string structure anywhere in the document
    emails = re.findall(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", full_text)
    
    contacts = []
    # Match them up by pair index
    for i in range(max(len(mobiles), len(emails))):
        mob = f"+65 {mobiles[i]}" if i < len(mobiles) else ""
        em = emails[i] if i < len(emails) else ""
        contacts.append(f"{mob} {em}".strip())
    contacts_field = "\n".join(contacts)
    
    # Extract Referral
    referral = "Google"
    if "referral" in full_text.lower():
        ref_segment = re.search(r"referral\s*,\s*(\w+)", full_text, re.IGNORECASE)
        if ref_segment and "jav" in ref_segment.group(1).lower():
            referral = "Javern"
            
    return matter_type, clients_field, contacts_field, referral

# --- 3. UI & RUNTIME LOGIC ---
uploaded_file = st.file_uploader("Drag and drop Open File Sheet (.docx) here", type=["docx"])

if uploaded_file is not None:
    try:
        with st.spinner("Processing data..."):
            sheet = get_google_sheet()
            
            # 1. Matter No Logic: Get all existing numbers from Column C to find max
            matter_nos = sheet.col_values(3)[1:]  
            valid_numbers = [int(val.strip()) for val in matter_nos if str(val).strip().isdigit()]
            new_matter_no = str(max(valid_numbers) + 1) if valid_numbers else "20260623"
            
            # 2. Find the REAL first empty row based on Column E (Clients) being blank
            client_values = sheet.col_values(5) # Column E
            
            # Find the first row index where Column E is empty (ignoring header row 1)
            target_row = 2
            while target_row <= len(client_values) and client_values[target_row - 1].strip() != "":
                target_row += 1
                
            # Calculate Serial Index for Column A based on the row number minus headers
            next_index = target_row - 1
            
            # 3. Dynamic layout updates
            today_date = datetime.now().strftime("%d %B %Y").lstrip("0")
            matter_type, clients, contacts, referral = extract_matter_data(uploaded_file)
            
            new_row = [
                next_index,       # Column A: Serial index
                today_date,       # Column B: Date Opened
                new_matter_no,    # Column C: Matter No
                matter_type,      # Column D: Type of Work
                clients,          # Column E: Client(s) Name
                contacts,         # Column F: Contact(s) Info
                referral,         # Column G: Referral Source
                "Yes"             # Column H: Bill Paid
            ]
            
            # Update the specific empty row directly (e.g., Row 29) instead of appending!
            # Format range like "A29:H29"
            cell_range = f"A{target_row}:H{target_row}"
            sheet.update(range_name=cell_range, values=[new_row])
            
        st.success(f"🚀 Success! Matter No {new_matter_no} written directly to Row {target_row}!")
        st.json({
            "Row Slot": target_row,
            "Matter No": new_matter_no,
            "Type": matter_type,
            "Clients": clients.split('\n'),
            "Contacts": contacts.split('\n')
        })

    except Exception as e:
        st.error(f"Error executing automation: {e}")