# EML to Excel Parser

This Streamlit application allows you to upload one or more `.eml` (email) files, parses key information from them, and provides a consolidated Excel (`.xlsx`) file for download.

## Features

*   **Multiple File Upload**: Upload one or more `.eml` files simultaneously.
*   **Key Information Extraction**: Parses the following details from each email:
    *   From
    *   To
    *   Subject
    *   Date
    *   Source IP (from 'Received' headers)
    *   SPF Result
    *   DKIM Results
    *   DMARC Result
    *   List-Unsubscribe Link
    *   Plain Text Body
    *   Message-ID
*   **Data Display**: Shows a combined table of all parsed email data within the app.
*   **Excel Export**: Download all parsed data as a single `.xlsx` file.

## Requirements

The application requires the following Python libraries:

*   streamlit
*   pandas
*   openpyxl

You can install them using the provided [`requirements.txt`](requirements.txt) file:

```bash
pip install -r requirements.txt
```

## How to Run

1.  **Install Dependencies**:
    ```bash
    pip install -r requirements.txt
    ```
2.  **Run the Streamlit App**:
    Navigate to the directory containing `better_parser.py` and run:
    ```bash
    streamlit run better_parser.py
    ```
3.  **Upload Files**:
    Open the URL provided by Streamlit (usually `http://localhost:8501`) in your web browser. Use the file uploader to select your `.eml` files.
4.  **Analyze and Download**:
    The app will process the files and display the extracted data. Click the "Download All as Excel (.xlsx)" button to save the results.

## File Structure

*   [`better_parser.py`](better_parser.py): The main Python script containing the Streamlit app logic and EML parsing functions.
*   [`requirements.txt`](requirements.txt): A list of Python dependencies for the project.
