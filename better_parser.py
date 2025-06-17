import streamlit as st
import email
from email import policy
from email.header import decode_header
from email.utils import parsedate_to_datetime
import re
import json
import base64
import pandas as pd
import io

# --- Helper Functions (largely the same as before) ---

def clean_header_value(value):
    """Decodes header values (like subject) and cleans them up."""
    if value is None:
        return None
    decoded_parts = decode_header(value)
    final_value = []
    for part, charset in decoded_parts:
        if isinstance(part, bytes):
            try:
                final_value.append(part.decode(charset or 'utf-8', errors='ignore'))
            except (UnicodeDecodeError, LookupError):
                final_value.append(part.decode('latin-1', errors='ignore'))
        else:
            final_value.append(str(part))
    return ''.join(final_value)

def extract_url_from_header(header_value):
    """Uses regex to find the first http/https URL within a header string."""
    if not header_value:
        return None
    match = re.search(r'<https?://[^>]+>', header_value)
    if match:
        return match.group(0).strip('<>')
    match = re.search(r'https?://\S+', header_value)
    if match:
        return match.group(0)
    return None

def parse_security_headers(auth_results_header):
    """Parses the Authentication-Results header for SPF, DKIM, and DMARC."""
    security_info = {
        "spf": {"result": "not_found", "domain": None},
        "dkim": [],
        "dmarc": {"result": "not_found", "domain": None}
    }
    if not auth_results_header:
        return security_info

    spf_match = re.search(r'spf=(\w+).*?smtp\.mailfrom=([\w\.-@]+)', auth_results_header)
    if spf_match:
        security_info["spf"]["result"] = spf_match.group(1)
        security_info["spf"]["domain"] = spf_match.group(2).split('@')[-1]

    dkim_matches = re.findall(r'dkim=(\w+)\s+header\.i=@([\w\.-]+)', auth_results_header)
    for match in dkim_matches:
        security_info["dkim"].append({"result": match[0], "domain": match[1]})
        
    dmarc_match = re.search(r'dmarc=(\w+).*?header\.from=([\w\.-]+)', auth_results_header)
    if dmarc_match:
        security_info["dmarc"]["result"] = dmarc_match.group(1)
        security_info["dmarc"]["domain"] = dmarc_match.group(2)

    return security_info

# --- Main Parsing and Data Transformation Logic ---

def parse_eml_to_dataframe(eml_bytes):
    """
    Parses EML bytes and transforms the data directly into a Pandas DataFrame.
    """
    msg = email.message_from_bytes(eml_bytes, policy=policy.default)

    # Decode core fields
    from_addr = clean_header_value(msg.get('From'))
    to_addrs = [addr for name, addr in email.utils.getaddresses(msg.get_all('To', []))]
    subject = clean_header_value(msg.get('Subject'))
    try:
        date_iso = parsedate_to_datetime(msg.get('Date')).isoformat()
    except (TypeError, ValueError):
        date_iso = None

    # Find the plain text body
    body_plain = ""
    if msg.is_multipart():
        for part in msg.walk():
            content_disposition = str(part.get("Content-Disposition"))
            if part.get_content_type() == "text/plain" and "attachment" not in content_disposition:
                payload = part.get_payload(decode=True)
                charset = part.get_content_charset() or 'utf-8'
                body_plain = payload.decode(charset, errors='replace').strip()
                break # Stop after finding the first plain text part
    else:
        if msg.get_content_type() == "text/plain":
            payload = msg.get_payload(decode=True)
            charset = msg.get_content_charset() or 'utf-8'
            body_plain = payload.decode(charset, errors='replace').strip()
    
    # Get security details
    auth_results_header = msg.get('Authentication-Results')
    security_details = parse_security_headers(auth_results_header)
    dkim_records = security_details.get('dkim', [])

    # Create a flattened dictionary for the DataFrame
    flat_data = {
        'from': email.utils.parseaddr(from_addr)[1],
        'to': ', '.join(to_addrs),
        'subject': subject,
        'date': date_iso,
        'source_ip': next((match.group(1) for h in msg.get_all('Received', []) if (match := re.search(r'\[(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})\]', h))), None),
        'spf_result': security_details.get('spf', {}).get('result'),
        'dkim_results': ', '.join([d.get('result', 'n/a') for d in dkim_records]),
        'dmarc_result': security_details.get('dmarc', {}).get('result'),
        'list_unsubscribe_link': extract_url_from_header(msg.get('List-Unsubscribe')),
        'body_plain': body_plain,
        'message_id': msg.get('Message-ID'),
    }

    # Create DataFrame
    df = pd.DataFrame([flat_data])
    return df

def convert_df_to_excel_bytes(df):
    """Converts a DataFrame to an in-memory Excel file (bytes)."""
    output_buffer = io.BytesIO()
    # Use the ExcelWriter to write the df to the buffer
    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='ParsedEmail')
    # Retrieve the bytes from the buffer
    excel_bytes = output_buffer.getvalue()
    return excel_bytes

# --- Streamlit UI ---

st.set_page_config(page_title="EML to Excel Parser", page_icon="📧", layout="wide")

st.title("📧 EML File to Excel Analyzer")
st.markdown("Upload one or more `.eml` files. The app will parse the key information and let you download it as a single `.xlsx` file.")

uploaded_files = st.file_uploader(
    "Choose your .eml files",
    type=['eml'],
    accept_multiple_files=True, # Allow multiple files
    help="Select one or more email files saved with the .eml extension."
)

if uploaded_files: # Check if the list is not empty
    all_dfs = [] # Initialize a list to store DataFrames from each file
    
    with st.spinner('Analyzing your email(s)...'):
        try:
            for uploaded_file in uploaded_files: # Iterate through each uploaded file
                st.write(f"Processing `{uploaded_file.name}`...")
                eml_bytes = uploaded_file.getvalue()
                
                # Parse the data directly into a DataFrame
                df = parse_eml_to_dataframe(eml_bytes)
                all_dfs.append(df) # Add the parsed DataFrame to our list
            
            if not all_dfs:
                st.warning("No data was parsed from the uploaded files.")
            else:
                # Concatenate all DataFrames into a single DataFrame
                combined_df = pd.concat(all_dfs, ignore_index=True)
                
                st.success(f"✅ {len(uploaded_files)} email(s) analyzed successfully!")
                
                # Display the combined DataFrame
                st.subheader("Parsed Email Data (Combined)")
                st.dataframe(combined_df)
                
                # Convert combined DataFrame to Excel bytes for download
                excel_data = convert_df_to_excel_bytes(combined_df)
                
                # Generate a dynamic filename
                download_filename = "parsed_emails.xlsx" # Generic name for multiple files
                
                st.download_button(
                    label="📥 Download All as Excel (.xlsx)",
                    data=excel_data,
                    file_name=download_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            
        except Exception as e:
            st.error(f"An error occurred during parsing: {e}")
            st.exception(e) # Provides a full traceback for debugging

else:
    st.info("Waiting for you to upload file(s).")

st.markdown("---")
st.write("Created with ❤️ by an AI Assistant")