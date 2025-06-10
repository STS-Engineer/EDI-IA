from flask import Flask, request, send_file, render_template_string
import pandas as pd
import os
import chardet
import time
from werkzeug.utils import secure_filename
from datetime import datetime, timedelta
from PyPDF2 import PdfReader
import re
import logging
logging.basicConfig(level=logging.INFO)
app = Flask(__name__)
from sqlalchemy import text
from sqlalchemy import create_engine

from sqlalchemy.engine import URL

db_url = URL.create(
    drivername="postgresql+psycopg2",
    username="adminavo",
    password="$#fKcdXPg4@ue8AW",  # No encoding needed
    host="avo-adb-001.postgres.database.azure.com",
    port=5432,
    database="EDI IA"
)
engine = create_engine(db_url)

# Directory for output files
OUTPUT_DIR = "outputs"
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)
 
ALLOWED_EXTENSIONS = {'csv', 'pdf', 'xls', 'xlsx'}
 

# HTML Template for Upload Form with Country Navigation
UPLOAD_FORM_HTML = """<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
    <title>CSV Transformer</title>
    <link rel="icon" href="/static/avo_carbon.jpg" type="image/jpg">
    <style>
      body {
        font-family: 'Arial', sans-serif;
        background-color: #f4f7f6;
        margin: 0;
        padding: 20px 0;
        min-height: 100vh;
      }
      .container {
        background-color: #ffffff;
        padding: 30px;
        border-radius: 10px;
        box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
        text-align: center;
        width: 90%;
        max-width: 1000px;
        margin: 0 auto;
        overflow: auto;
        max-height: 90vh;
      }
      
      /* Navigation Tabs */
      .nav-tabs {
        display: flex;
        border-bottom: 3px solid #ddd;
        margin-bottom: 30px;
        justify-content: center;
      }
      .nav-tab {
        background-color: #f8f9fa;
        border: 1px solid #ddd;
        border-bottom: none;
        padding: 15px 30px;
        margin-right: 5px;
        cursor: pointer;
        font-size: 16px;
        font-weight: bold;
        text-decoration: none;
        color: #555;
        border-radius: 8px 8px 0 0;
        transition: all 0.3s ease;
        display: flex;
        align-items: center;
        gap: 10px;
      }
      .nav-tab:hover {
        background-color: #e9ecef;
        color: #333;
      }
      .nav-tab.active {
        background-color: #28a745;
        color: white;
        border-color: #28a745;
      }
      .nav-tab::before {
        content: '';
        width: 20px;
        height: 15px;
        background-size: contain;
        background-repeat: no-repeat;
      }
      .nav-tab.germany::before {
        background-image: url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 5 3"><rect width="5" height="1" fill="%23000"/><rect width="5" height="1" y="1" fill="%23dd0000"/><rect width="5" height="1" y="2" fill="%23ffce00"/></svg>');
      }
      .nav-tab.tunisia::before {
        background-image: url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 3 2"><rect width="3" height="2" fill="%23e70013"/><circle cx="1" cy="1" r="0.4" fill="white"/><circle cx="1" cy="1" r="0.3" fill="%23e70013"/><polygon points="1.1,0.8 1.2,1 1.4,1 1.25,1.1 1.3,1.3 1.1,1.2 0.9,1.3 0.95,1.1 0.8,1 1,1" fill="white"/></svg>');
      }
      
      /* Tab Content */
      .tab-content {
        display: none;
      }
      .tab-content.active {
        display: block;
      }
      
      h1 {
        color: #333;
        font-size: 24px;
        margin-bottom: 20px;
      }
      h2 {
        color: #28a745;
        font-size: 20px;
        margin: 30px 0 20px 0;
      }
      p {
        font-size: 16px;
        color: #555;
        margin-bottom: 20px;
      }
      form {
        display: flex;
        flex-direction: column;
        align-items: center;
        gap: 15px;
      }
      input[type="file"], input[type="text"], input[type="number"], input[type="submit"], a.download-btn {
        padding: 12px;
        font-size: 16px;
        border: 1px solid #ddd;
        border-radius: 5px;
        outline: none;
        transition: all 0.3s ease-in-out;
        width: 100%;
        max-width: 400px;
        margin: 0 auto;
      }
      input[type="file"], input[type="text"], input[type="number"] {
        background-color: #f8f9fa;
      }
      input[type="file"]:hover, input[type="text"]:hover, input[type="number"]:hover {
        background-color: #e9ecef;
      }
      input[type="submit"], a.download-btn {
        background-color: #28a745;
        color: white;
        cursor: pointer;
        border: none;
        text-decoration: none;
        display: inline-block;
        text-align: center;
      }
      input[type="submit"]:hover, a.download-btn:hover {
        background-color: #218838;
      }
      table {
        width: 100%;
        border-collapse: collapse;
        margin-top: 20px;
      }
      table th, table td {
        border: 1px solid #ddd;
        padding: 8px;
        text-align: left;
      }
      table th {
        background-color: #28a745;
        color: white;
      }
      table tbody tr:nth-child(odd) {
        background-color: #f9f9f9;
      }
      table tbody tr:hover {
        background-color: #e9ecef;
      }
      .scrollable {
        overflow-x: auto;
        margin-top: 20px;
        max-height: 300px;
      }
      footer {
        margin-top: 40px;
        color: #888;
        font-size: 14px;
        border-top: 1px solid #eee;
        padding-top: 20px;
      }
      img {
        max-width: 300px;
        margin-bottom: 20px;
      }
      select.select-btn {
        appearance: none;
        background-color: #007bff;
        color: white;
        font-size: 16px;
        padding: 12px;
        border: none;
        border-radius: 5px;
        cursor: pointer;
        text-align: center;
        width: 100%;
        max-width: 400px;
        margin: 10px auto;
      }
      select.select-btn:hover {
        background-color: #0056b3;
      }
      .error-message {
        color: #dc3545;
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        padding: 10px;
        border-radius: 5px;
        margin: 10px 0;
      }
      .success-message {
        color: #155724;
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        padding: 10px;
        border-radius: 5px;
        margin: 10px 0;
      }
    </style>
 
  </head>
  <body>
    <div class="container">
      <img src="/static/logo-avocarbon.png" alt="Logo">
      
      <!-- Navigation Tabs -->
      <div class="nav-tabs">
        <a href="#" class="nav-tab germany" onclick="showTab('germany')" id="germany-tab">
          Germany - CSV Transformation
        </a>
        <a href="#" class="nav-tab tunisia" onclick="showTab('tunisia')" id="tunisia-tab">
          Tunisia - Excel Database Import
        </a>
      </div>

      <!-- Germany Tab Content - CSV Transformation -->
      <div id="germany-content" class="tab-content">
        <!-- Germany Tab Messages -->
        {% if germany_error_message %}
        <div class="error-message">{{ germany_error_message }}</div>
        {% endif %}
        {% if germany_success_message %}
        <div class="success-message">{{ germany_success_message }}</div>
        {% endif %}
        <h1>CSV & PDF Transformation Tool</h1>
        <p>Upload your CSV file or PDF file and select the customer to transform your data.</p>
        
        <form action="/convert" method="post" enctype="multipart/form-data">
          <input type="file" name="file" accept=".csv,.pdf" required>
          <select name="customer_name" class="btn select-btn" id="customer_name" required>
            <option value="" disabled selected>Select Customer Name</option>
            <option value="VALEO Poland">VALEO Poland</option>
            <option value="Bosch Brasil">Bosch Brasil</option>
            <option value="Bosch Budweis">Bosch Budweis</option>
            <option value="Bosch China">Bosch China</option>
            <option value="Nidec Poland">Nidec Poland</option>
          </select>
          <input type="submit" value="Transform Data">
        </form>

        {% if transformed_table %}
        <div class="scrollable">
          <h2>Transformed Data Preview</h2>
          <table>{{ transformed_table|safe }}</table>
        </div>
        {% endif %}

       {% if show_download %}
        <a href="/download/{{ download_path }}" class="download-btn">Download Transformed CSV</a>
        <p><strong>File saved as:</strong> {{ download_path }}</p>
        <form action="/send_to_db" method="post">
            <input type="hidden" name="csv_file" value="{{ download_path }}">
            <input type="hidden" name="customer_name" value="{{ request.form.get('customer_name', '') }}">
            <input type="date" name="file_date" placeholder="Select File Date" required style="margin-bottom: 10px;">
            <input type="submit" value="Send to Database" class="download-btn" style="background-color: #007bff;">
        </form>
        {% endif %}
      </div>

       <!-- Tunisia Tab Content - Excel Database Import -->
      <div id="tunisia-content" class="tab-content">
        <!-- Tunisia Tab Messages -->
        {% if tunisia_error_message %}
        <div class="error-message">{{ tunisia_error_message }}</div>
        {% endif %}
        {% if tunisia_success_message %}
        <div class="success-message">{{ tunisia_success_message }}</div>
        {% endif %}
        <h1>Excel Database Import Tool</h1>
        <p>Upload an Excel file (.xls or .xlsx) to send rows to the PostgreSQL database table <strong>EDI</strong>.</p>

        <form action="/preview_excel" method="post" enctype="multipart/form-data">
          <input type="file" name="excel_file" accept=".xls,.xlsx" required>
          <input type="submit" value="Preview Excel Data">
        </form>

        {% if excel_table %}
        <div class="scrollable">
          <h2>Excel Data Preview</h2>
          <table>{{ excel_table|safe }}</table>
        </div>

        <form action="/insert_excel" method="post">
          <input type="hidden" name="temp_file" value="{{ temp_file }}">
          <input type="number" name="week_number" placeholder="Enter Week Number" required>
          <input type="submit" value="Send to Database">
        </form>
        {% endif %}
      </div>

      <footer>
        &copy; 2025 Data Processing Tool. All rights reserved. Powered by STS
      </footer>
    </div>

    <script>
      function showTab(country) {
        // Hide all tab contents
        const contents = document.querySelectorAll('.tab-content');
        contents.forEach(content => content.classList.remove('active'));
        
        // Remove active class from all tabs
        const tabs = document.querySelectorAll('.nav-tab');
        tabs.forEach(tab => tab.classList.remove('active'));
        
        // Show selected tab content
        document.getElementById(country + '-content').classList.add('active');
        document.getElementById(country + '-tab').classList.add('active');
        
        // Store the active tab in localStorage
        localStorage.setItem('activeTab', country);
      }
      
      // Initialize the page
      document.addEventListener('DOMContentLoaded', function() {
        // Check if there's a tab preference from form submission
        const urlParams = new URLSearchParams(window.location.search);
        const formTab = urlParams.get('tab');
        
        // Get the active tab from form submission, localStorage, or default to germany
        const activeTab = formTab || localStorage.getItem('activeTab') || 'germany';
        showTab(activeTab);
      });
    </script>
  </body>
</html>
"""
 
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS
 
def detect_encoding(file):
    raw_data = file.read()
    result = chardet.detect(raw_data)
    file.seek(0)
    return result['encoding']
 
def extract_before_date(line):
    """
    Extracts specific values (FORECAST, Backlog, Firm) before a date in a line of text.
    If none of these are found, extracts the first matching value before the date.
    """
    pattern = r"(.*?)\s+(\d{2}[-/]\d{2}[-/]\d{4}|\d{4}[-/]\d{2}[-/]\d{2})"
    match = re.search(pattern, line)
 
    if match:
        before_date = match.group(1).strip()
        keywords = ["FORECAST", "Backlog", "Firm"]
        for keyword in keywords:
            if keyword.lower() in before_date.lower():
                return keyword
        return before_date
    return None
 
def extract_date_and_number(text):
    date_number_pattern = r"(\d{2}[-/]\d{2}[-/]\d{4}|\d{4}[-/]\d{2}[-/]\d{2})\s+(\d+(?:,\d{3})*)"
    matches = re.findall(date_number_pattern, text)
    cleaned_matches = [(date, int(number.replace(",", ""))) for date, number in matches]
    return cleaned_matches
 
def parse_pdf(file):
    temp_filename = secure_filename(file.filename)
    temp_filepath = os.path.join(OUTPUT_DIR, temp_filename)
   
    file.save(temp_filepath)
   
    try:
        with open(temp_filepath, 'rb') as pdf_file:
            reader = PdfReader(pdf_file)
            text = ''
            for page in reader.pages:
                text += page.extract_text() + '\n'
    finally:
        # Always remove temp file, even if there's an error
        if os.path.exists(temp_filepath):
            os.remove(temp_filepath)
   
    return text
 
# Customer configuration mapping
CUSTOMER_CONFIG = {
    'VALEO Poland': {
        'code': 'VPL',
        'processor': 'valeo'
    },
    'Bosch Poland': {
        'code': 'BPL',
        'processor': 'bosch'
    },
    'Bosch Budweis': {
        'code': 'BBW',
        'processor': 'bosch'
    },
      'Bosch Brasil': {
        'code': 'BB',
        'processor': 'bosch'
    },
    'Bosch China': {
        'code': 'BCN',
        'processor': 'bosch'
    },
    'Nidec Poland': {
        'code': 'NPL',
        'processor': 'nidec'
    }
}
 
def process_csv(file, customer_name):
    """Process CSV file and return transformed DataFrame and filename"""
    try:
        # Get customer configuration
        if customer_name not in CUSTOMER_CONFIG:
            raise ValueError(f"Unknown customer: {customer_name}")
       
        config = CUSTOMER_CONFIG[customer_name]
        customer_code = config['code']
       
        # Detect encoding
        encoding = detect_encoding(file)
       
        # Read CSV
        df = pd.read_csv(file, encoding=encoding)
       
        # Basic transformation
        df['TIERSLU'] = customer_code
        df['Libelle client'] = customer_name
        df['Tiers livré'] = customer_code
       
        # Save transformed data
        output_filename = f"{customer_code}_{int(time.time())}.csv"
        output_path = os.path.join(OUTPUT_DIR, output_filename)
        df.to_csv(output_path, index=False, sep=';')
       
        logging.info(f"CSV data saved to {output_path}")
        return df, output_filename
       
    except Exception as e:
        logging.error(f"Error processing CSV: {e}")
        raise
 

def store_to_postgres_edi(df):
    try:
        with engine.begin() as conn:
            for _, row in df.iterrows():
                # Handle ExpectedDeliveryDate - it should always be present now
                expected_date = row.get("ExpectedDeliveryDate")
                if pd.isna(expected_date) or expected_date is None or str(expected_date).strip() in ['', 'nan', 'None']:
                    raise ValueError(f"ExpectedDeliveryDate is null for row with ClientCode: {row['ClientCode']}, ProductCode: {row['ProductCode']}. This should not happen if the Excel file contains 'Livraison au plus tard' data.")
                
                stmt = text("""
                    INSERT INTO "EDITunisia" (
                        "ClientCode", "ProductCode", "Date", "Quantity", "EDIWeekNumber", "ExpectedDeliveryDate"
                    )
                    VALUES (:client, :product, :date, :qty, :week, :expected)
                    ON CONFLICT ("ClientCode", "ProductCode", "Date", "EDIWeekNumber", "ExpectedDeliveryDate") DO NOTHING
                """)
                conn.execute(stmt, {
                    "client": str(row["ClientCode"]),
                    "product": str(row["ProductCode"]),
                    "date": str(row["Date"]),
                    "qty": int(row["Quantity"]),
                    "week": int(row["EDIWeekNumber"]),
                    "expected": str(expected_date)
                })
        logging.info(f"Insert attempted for {len(df)} rows (duplicates skipped).")
    except Exception as e:
        logging.error(f"Error inserting into EDITunisia table: {e}")
        raise






def process_excel(file, customer_name):
    try:
        # Read Excel file with proper encoding handling
        df = pd.read_excel(file, engine='openpyxl')
        
        # Print column names for debugging
        print("Available columns:", df.columns.tolist())
        
        # Clean column names - remove extra spaces and handle encoding issues
        df.columns = df.columns.str.strip()
        
        # Map the actual French column names to expected English names
        # Handle both encoded and properly decoded versions
        column_mapping = {
            "Numéro Client": "ClientCode",
            "NumÃ©ro Client": "ClientCode",  # Encoded version
            "Code article": "ProductCode",
            "Q.Cadence prévue": "Quantity",
            "Q.Cadence prÃ©vue": "Quantity",  # Encoded version
            "Dépôt": "Date",  # Add the depot/date column
            "DÃ©pÃ´t": "Date",  # Encoded version of Dépôt
            "Livraison au plus tard": "ExpectedDeliveryDate",
            "Livraison au plus tard": "ExpectedDeliveryDate",  # Properly decoded version
            "Livraison au plus tard": "ExpectedDeliveryDate"   # Encoded version if any
        }
        
        # Apply the mapping
        df = df.rename(columns=column_mapping)
        
        # Check if required columns exist after mapping
        required_columns = ["ClientCode", "ProductCode", "Quantity", "Date", "ExpectedDeliveryDate"]
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            # Try alternative column names if standard mapping fails
            alt_mapping = {}
            for col in df.columns:
                if "client" in col.lower() or "numéro" in col.lower():
                    alt_mapping[col] = "ClientCode"
                elif "code" in col.lower() and "article" in col.lower():
                    alt_mapping[col] = "ProductCode"
                elif "cadence" in col.lower() or "quantité" in col.lower():
                    alt_mapping[col] = "Quantity"
                elif "dépôt" in col.lower() or "depot" in col.lower():
                    alt_mapping[col] = "Date"
                elif "livraison" in col.lower() and ("tard" in col.lower() or "plus" in col.lower()):
                    alt_mapping[col] = "ExpectedDeliveryDate"
            
            df = df.rename(columns=alt_mapping)
            
            # Check again after alternative mapping
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                raise ValueError(f"Required columns missing after mapping: {missing_columns}. Available columns: {df.columns.tolist()}")

        # Drop rows with missing required fields
        df = df.dropna(subset=required_columns)
        
        if df.empty:
            raise ValueError("No valid data rows found after removing empty values")

        # Convert to correct types with error handling
        try:
            # Keep ClientCode and ProductCode as strings since they contain alphanumeric values
            df["ClientCode"] = df["ClientCode"].astype(str).str.strip()
            df["ProductCode"] = df["ProductCode"].astype(str).str.strip()
            
            # Convert Date column to string and clean it
            df["Date"] = df["Date"].astype(str).str.strip()
            
            # Handle ExpectedDeliveryDate - this should come from "Livraison au plus tard"
            if "ExpectedDeliveryDate" in df.columns:
                # Convert to string and clean up
                df["ExpectedDeliveryDate"] = df["ExpectedDeliveryDate"].astype(str).str.strip()
                # Check for truly empty values and raise error if all are empty
                empty_mask = df["ExpectedDeliveryDate"].isin(['', 'nan', 'None', 'NaT']) | df["ExpectedDeliveryDate"].isna()
                if empty_mask.all():
                    raise ValueError("ExpectedDeliveryDate column exists but all values are empty. Please check your 'Livraison au plus tard' column in the Excel file.")
                # For individual empty values, we could use Date as fallback or raise an error
                if empty_mask.any():
                    print(f"Warning: {empty_mask.sum()} rows have empty ExpectedDeliveryDate values")
                    # Option 1: Use Date as fallback for empty values
                    df.loc[empty_mask, "ExpectedDeliveryDate"] = df.loc[empty_mask, "Date"]
                    # Option 2: Remove rows with empty ExpectedDeliveryDate
                    # df = df[~empty_mask]
            else:
                raise ValueError("ExpectedDeliveryDate column not found. Please ensure your Excel file has a 'Livraison au plus tard' column.")
            
            # Only convert Quantity to integer
            df["Quantity"] = pd.to_numeric(df["Quantity"], errors='coerce').astype('Int64')
        except Exception as e:
            raise ValueError(f"Error converting data types: {e}")
        
        # Remove rows where Quantity conversion failed (NaN values)
        df = df.dropna(subset=["Quantity"])
        
        if df.empty:
            raise ValueError("No valid data rows found after type conversion")

        # Keep only necessary columns in correct order
        df = df[["ClientCode", "ProductCode", "Date", "Quantity", "ExpectedDeliveryDate"]]

        # Insert into PostgreSQL
        store_to_postgres_edi(df)

        # Save preview CSV
        output_filename = f"EDITunisia_{int(time.time())}.csv"
        output_path = os.path.join(OUTPUT_DIR, output_filename)
        df.to_csv(output_path, index=False, sep=';')

        logging.info(f"Excel data saved to {output_path} and inserted into database.")
        return df, output_filename

    except Exception as e:
        logging.error(f"Error processing Excel: {e}")
        raise



def process_valeo_pdf(file, customer_code, customer_name, output_dir=None):
    """Process VALEO PDF format - Updated with material mapping and simplified output"""
    if output_dir is None:
        output_dir = OUTPUT_DIR

    try:
        text = parse_pdf(file)
    except Exception as e:
        logging.error(f"Failed to parse PDF file: {e}")
        raise ValueError(f"Failed to parse PDF file: {e}")

    # Material mappings
    material_map = {
        "1023093": "190313",
        "1023645": "191663",
        "1026188": "187144",
        "1026258": "194470",
        "1026540": "202066",
        "1026629": "214188"
    }
    reverse_map = {v: k for k, v in material_map.items()}

    sections = re.split(r"Sachnummer:\s+", text)[1:]
    results = []

    for section in sections:
        
         # Clean out metadata that can interfere with parsing
        section = re.sub(r"Druckdatum:\s*\d{2}\.\d{2}\.\d{4}\s+\d{2}:\d{2}", "", section)
        section = re.sub(r"Seite:\s*\d+\s+von\s+\d+", "", section)
        # Extract client name
        client = customer_name.split(" C")[0]

        # Extract material number from start of section and clean it
        material_match = re.match(r"0*(\d+)", section)
        raw_client_code = material_match.group(1) if material_match else None

        # Extract article (Bezeichnung)
        article_match = re.search(r"Bezeichnung:\s*(.*?)\s+Ersetzt", section)
        article = article_match.group(1).strip() if article_match else "UNKNOWN"

        # Extract material description from 'Materialbeschreibung (Kunde):'
        mat_desc_match = re.search(r"Materialbeschreibung\s+\(Kunde\):\s*(.*?)\s+(?:Ersetzt|Sachnr\.|Liefer|\d{2}\.\d{2}\.\d{4})", section, re.DOTALL)
        material_description = mat_desc_match.group(1).strip() if mat_desc_match else article

        # Determine AVO and client codes via mapping
        avo_code = customer_code_val = "UNKNOWN"
        if raw_client_code:
            if raw_client_code in material_map:
                avo_code = raw_client_code
                customer_code_val = material_map[raw_client_code]
            elif raw_client_code in reverse_map:
                customer_code_val = raw_client_code
                avo_code = reverse_map[raw_client_code]

        # Extract delivery dates
        date_qty = re.findall(r"(\d{2}\.\d{2}\.\d{4})\s+(\d+)", section)
        for date_str, qty_str in date_qty:
            if int(qty_str) < 100:  # ignore time values like 11:59
                continue
            results.append({
                    "Client name": client,
                    "Material Description": material_description,
                    "Client Material No": customer_code_val,
                    "AVO Material No": avo_code,
                    "Quantity": int(qty_str),
                    "Date from": date_str,
                    "Date until": ""
                })

        # Extract week ranges
        week_ranges = re.findall(r"(20\d{2}\s+w\d{2})\s*-\s*(20\d{2}\s+w\d{2})\s+(\d+)", section)
        for from_week, to_week, qty_str in week_ranges:
            results.append({
                "Client name": client,
                "Material Description": material_description,
                "Client Material No": customer_code_val,
                "AVO Material No": avo_code,
                "Quantity": int(qty_str),
                "Date from": from_week.strip(),
                "Date until": to_week.strip()
            })

    if not results:
        raise ValueError("No valid delivery data found in the VALEO PDF.")

    df = pd.DataFrame(results)

    timestamp = int(time.time())
    output_filename = f"{customer_code}_{timestamp}.csv"
    output_path = os.path.join(output_dir, output_filename)

    os.makedirs(output_dir, exist_ok=True)
    df.to_csv(output_path, index=False, sep=';')
    logging.info(f"VALEO PDF data saved to {output_path}")

    return df, output_filename



def process_nidec_pdf(file, customer_code, customer_name):
    """Process NIDEC structured call-off PDF format with material mapping."""
    try:
        text = parse_pdf(file)

        # Material mappings for Nidec Poland
        nidec_material_map = {
            "1022201": "471-695-99-99",
            "1027700": "503-660-99-99"
        }
        nidec_reverse_map = {v: k for k, v in nidec_material_map.items()}

        # Extract static fields
        material_match = re.search(r"Material.*?(\d{3}-\d{3}-\d{2}-\d{2})", text, re.DOTALL)
        description_match = re.search(r"Material description\s+([A-Z \-/]+)", text)

        material = material_match.group(1).strip() if material_match else "UNKNOWN"
        material_description = description_match.group(1).strip() if description_match else "UNKNOWN"

        # Remove any trailing 'C' character from material description if it exists
        if material_description.endswith("C"):
            material_description = material_description[:-1].strip()

        # Resolve material mapping
        avo_code = client_code = "UNKNOWN"
        if material in nidec_reverse_map:
            avo_code = nidec_reverse_map[material]
            client_code = material
        elif material in nidec_material_map:
            avo_code = material
            client_code = nidec_material_map[material]

        # Match the planning table lines with optional Date until
        pattern = re.compile(r"(\d{2}\.\d{2}\.\d{4})(?:\s+(\d{2}\.\d{2}\.\d{4}))?\s+([\d.,]+)\s+PCE\s+([\d.,]+)")

        results = []
        for match in pattern.findall(text):
            date_from, date_until, despatch_qty, cum_qty = match

            results.append({
                "Client name": customer_name,
                "Client Material No": client_code,
                "AVO Material No": avo_code,
                "Material Description": material_description,
                "Date from": date_from,
                "Date until": date_until.strip() if date_until else "",
                "Quantity": despatch_qty.replace(".", "").replace(",", "."),
                "Unit": "PCE",
                "Cum. quantity": cum_qty.replace(".", "").replace(",", "."),
                
            })

        if not results:
            raise ValueError("No delivery entries found in the NIDEC PDF.")

        df = pd.DataFrame(results)

        output_filename = f"{customer_code}_{int(time.time())}_Calloff.csv"
        output_path = os.path.join(OUTPUT_DIR, output_filename)
        df.to_csv(output_path, index=False, sep=';')

        logging.info(f"NIDEC structured PDF data saved to {output_path}")
        return df, output_filename

    except Exception as e:
        logging.error(f"Error processing NIDEC PDF: {e}")
        raise



def process_bosch_pdf(file, customer_code, customer_name, output_dir=None):
    """Process BOSCH PDF format - Fixed version with correct column parsing and number conversion"""
    if output_dir is None:
        output_dir = OUTPUT_DIR

    text = parse_pdf(file)
    lines = text.splitlines()

    organisation = "Robert Bosch, spol. s r. o."
    material = ""
    druckdatum = ""
    material_description = "UNKNOWN"

    # Material mappings
    bosch_material_map = {
        "1027599": "1582875601",
        "1022031": "1582884102",
        "1026644": "1394320515",
        "1021731": "1394320230"
    }
    bosch_reverse_map = {v: k for k, v in bosch_material_map.items()}

    # Function to convert European number format to integer
    def convert_european_number(value):
        """Convert European number format to integer"""
        if not value or pd.isna(value):
            return 0
        
        # Convert to string if not already
        str_value = str(value).strip()
        
        # Handle European format: dots as thousand separators, comma as decimal
        if ',' in str_value:
            # Has decimal part - split and handle
            parts = str_value.split(',')
            integer_part = parts[0].replace('.', '')  # Remove thousand separators
            # For integer conversion, we ignore decimal part
            try:
                return int(integer_part)
            except ValueError:
                return 0
        else:
            # No comma, so dots are thousand separators
            try:
                return int(str_value.replace('.', ''))
            except ValueError:
                return 0

    for line in lines:
        if "Organisation:" in line:
            organisation = line.split("Organisation:")[-1].strip()
        if "Material:" in line:
            match = re.search(r"Material:\s*(\S+)", line)
            if match:
                material = match.group(1)
        if "Druckdatum:" in line:
            match = re.search(r"Druckdatum:\s*(\d{2}\.\d{2}\.\d{2})", line)
            if match:
                druckdatum = match.group(1)
        if "Materialbeschreibung (Kunde):" in line:
            match = re.search(r"Materialbeschreibung \(Kunde\):\s*(.*)", line)
            if match:
                material_description = match.group(1).strip()

    # Resolve material mapping
    avo_code = "UNKNOWN"
    client_code = material  # Always use the found material as client material number
    if material in bosch_material_map:
        avo_code = material
        client_code = bosch_material_map[material]
    elif material in bosch_reverse_map:
        avo_code = bosch_reverse_map[material]
        client_code = material

    # Updated regex patterns to handle different Bosch PDF formats
    # Pattern 1: Full format with times: Liefertermin Time Abholtermin Time Liefermenge EFZ Verbindlichkeitsstufe
    pattern_with_times = re.compile(
        r"(\d{2}\.\d{2}\.\d{2})\s+(\d{2}:\d{2})\s+"  # Liefertermin + Time
        r"(\d{2}\.\d{2}\.\d{2})\s+(\d{2}:\d{2})\s+"  # Abholtermin + Time
        r"([\d.,]+)\s+"                              # Liefermenge
        r"([\d.,]+)\s+"                              # EFZ
        r"(\w+)",                                    # Verbindlichkeitsstufe
        re.UNICODE
    )

    # Pattern 2: Format without times: Liefertermin Abholtermin Liefermenge EFZ Differenz Verbindlichkeitsstufe
    pattern_no_times = re.compile(
        r"(\d{2}\.\d{2}\.\d{2})\s+"  # Liefertermin
        r"(\d{2}\.\d{2}\.\d{2})\s+"  # Abholtermin
        r"([\d.,]+)\s+"              # Liefermenge
        r"([\d.,]+)\s+"              # EFZ
        r"([\d.,]+)\s+"              # Differenz (optional)
        r"(\w+)",                    # Verbindlichkeitsstufe
        re.UNICODE
    )

    # Pattern 3: Short format (only pickup date): Abholtermin Liefermenge EFZ Verbindlichkeitsstufe
    pattern_short = re.compile(
        r"(\d{2}\.\d{2}\.\d{2})\s+"  # Abholtermin (pickup date only)
        r"([\d.,]+)\s+"              # Liefermenge
        r"([\d.,]+)\s+"              # EFZ
        r"(\w+)",                    # Verbindlichkeitsstufe
        re.UNICODE
    )

    data = []

    for line in lines:
        line_stripped = line.strip()
        
        # Try pattern with times first (most complete format)
        match_with_times = pattern_with_times.match(line_stripped)
        if match_with_times:
            liefertermin, lieferzeit, abholtermin, abholzeit, liefermenge, efz, stufe = match_with_times.groups()
            # Combine date and time for display
            liefertermin_full = liefertermin
            abholtermin_full = abholtermin
        else:
            # Try pattern without times
            match_no_times = pattern_no_times.match(line_stripped)
            if match_no_times:
                liefertermin, abholtermin, liefermenge, efz, differenz, stufe = match_no_times.groups()
                liefertermin_full = liefertermin
                abholtermin_full = abholtermin
            else:
                # Try short pattern (only pickup date)
                match_short = pattern_short.match(line_stripped)
                if match_short:
                    abholtermin, liefermenge, efz, stufe = match_short.groups()
                    liefertermin_full = ""  # No delivery date in this format
                    abholtermin_full = abholtermin
                else:
                    continue

        # Convert European number format to integers
        liefermenge_converted = convert_european_number(liefermenge)
        efz_converted = convert_european_number(efz)

        # Calculate difference
        try:
            previous_efz = float(data[-1]['EFZ_calc']) if data else 0
            differenz_calc = efz_converted - previous_efz
        except Exception:
            differenz_calc = 0

        data.append({
            "Client name": customer_name,
            "Delivery date": liefertermin_full,
            "Pickup date": abholtermin_full,
            "Quantity": liefermenge_converted,  # Now converted to integer
            "EFZ": efz_converted,  # Now converted to integer
            "EFZ_calc": efz_converted,  # For calculation purposes
            "Commitment level": stufe,
            "AVO Material No": avo_code,
            "Client Material No": client_code,
            "Material Description": material_description
        })

    if not data:
        raise ValueError("Keine Planungsdaten aus der BOSCH PDF extrahiert.")

    df = pd.DataFrame(data)
    
    # Remove the calculation helper column before saving
    if 'EFZ_calc' in df.columns:
        df = df.drop('EFZ_calc', axis=1)

    # Create output filename with customer code and timestamp
    output_filename = f"{customer_code}_{int(time.time())}_Plan.csv"
    output_path = os.path.join(output_dir, output_filename)
    df.to_csv(output_path, index=False, sep=';')

    logging.info(f"BOSCH PDF data saved to {output_path}")
    return df, output_filename

 
 
def process_pdf_by_company(file, customer_name):
    """Route PDF processing to the appropriate company processor"""
    try:
        if customer_name not in CUSTOMER_CONFIG:
            raise ValueError(f"Unknown customer: {customer_name}")
       
        config = CUSTOMER_CONFIG[customer_name]
        customer_code = config['code']
        processor = config['processor']
       
        if processor == 'valeo':
            return process_valeo_pdf(file, customer_code, customer_name)
           
        elif processor == 'bosch':
            return process_bosch_pdf(file, customer_code, customer_name)
           
        elif processor == 'nidec':
            return process_nidec_pdf(file, customer_code, customer_name)
           
        else:
            raise ValueError(f"Unknown processor type: {processor}")
           
    except Exception as e:
        logging.error(f"Error processing PDF for {customer_name}: {e}")
        raise
 
@app.route('/')
def index():
    return render_template_string(UPLOAD_FORM_HTML,
                                summary_table=None,
                                show_download=False,
                                germany_error_message=None,
                                germany_success_message=None,
                                tunisia_error_message=None,
                                tunisia_success_message=None)
 
@app.route('/convert', methods=['POST'])
def convert():
    try:
        if 'file' not in request.files or 'customer_name' not in request.form:
            return render_template_string(UPLOAD_FORM_HTML,
                                        summary_table=None,
                                        show_download=False,
                                        germany_error_message="Missing file or customer name",
                                        germany_success_message=None,
                                        tunisia_error_message=None,
                                        tunisia_success_message=None)
 
        file = request.files['file']
        customer_name = request.form['customer_name']
 
        if not file.filename:
            return render_template_string(UPLOAD_FORM_HTML,
                                        summary_table=None,
                                        show_download=False,
                                        germany_error_message="No file selected",
                                        germany_success_message=None,
                                        tunisia_error_message=None,
                                        tunisia_success_message=None)
 
        if not allowed_file(file.filename):
            return render_template_string(UPLOAD_FORM_HTML,
                                        summary_table=None,
                                        show_download=False,
                                        germany_error_message="Invalid file type. Please upload a CSV or PDF file.",
                                        germany_success_message=None,
                                        tunisia_error_message=None,
                                        tunisia_success_message=None)
 
        # Process file based on type
        if file.filename.lower().endswith('.pdf'):
            transformed_data, output_filename = process_pdf_by_company(file, customer_name)
        elif file.filename.lower().endswith('.csv'):
            transformed_data, output_filename = process_csv(file, customer_name)
        elif file.filename.lower().endswith(('.xls', '.xlsx')):
            transformed_data, output_filename = process_excel(file, customer_name)
        else:
            return render_template_string(UPLOAD_FORM_HTML,
                                        summary_table=None,
                                        show_download=False,
                                        germany_error_message="Invalid file type. Please upload a CSV or PDF file.",
                                        germany_success_message=None,
                                        tunisia_error_message=None,
                                        tunisia_success_message=None)
 
        # Verify the file was created
        output_path = os.path.join(OUTPUT_DIR, output_filename)
        if not os.path.exists(output_path):
            return render_template_string(UPLOAD_FORM_HTML,
                                        summary_table=None,
                                        show_download=False,
                                        germany_error_message=f"File creation failed. Expected file: {output_path}",
                                        germany_success_message=None,
                                        tunisia_error_message=None,
                                        tunisia_success_message=None)
 
        # Display preview (limit to first 20 rows for performance)
        preview_data = transformed_data.head(20)
        transformed_table_html = preview_data.to_html(classes='table', index=False)
 
        return render_template_string(
            UPLOAD_FORM_HTML,
            transformed_table=transformed_table_html,
            excel_table=None,
            show_download=True,
            download_path=output_filename,
            germany_success_message=f"File processed successfully! {len(transformed_data)} rows transformed.",
            germany_error_message=None,
            tunisia_error_message=None,
            tunisia_success_message=None
        )
 
    except Exception as e:
        logging.error(f"Error in convert route: {e}")
        return render_template_string(UPLOAD_FORM_HTML,
                                    summary_table=None,
                                    show_download=False,
                                    germany_error_message=f"Error processing the file: {str(e)}",
                                    germany_success_message=None,
                                    tunisia_error_message=None,
                                    tunisia_success_message=None)

@app.route('/preview_excel', methods=['POST'])
def preview_excel():
    file = request.files.get('excel_file')
    if not file or not allowed_file(file.filename):
        return render_template_string(UPLOAD_FORM_HTML,
                                      germany_error_message=None,
                                      germany_success_message=None,
                                      tunisia_error_message="Please upload a valid Excel file.",
                                      tunisia_success_message=None)
    temp_path = os.path.join(OUTPUT_DIR, f"preview_{int(time.time())}.xlsx")
    file.save(temp_path)

    try:
        # Read Excel file with proper encoding
        df = pd.read_excel(temp_path, engine='openpyxl')
        
        # Clean column names
        df.columns = df.columns.str.strip()
        
        # Map French column names to English
        column_mapping = {
            "Numéro Client": "ClientCode",
            "NumÃ©ro Client": "ClientCode",
            "Code article": "ProductCode", 
            "Q.Cadence prévue": "Quantity",
            "Q.Cadence prÃ©vue": "Quantity",
            "Dépôt": "Date",  # Add the depot/date column
            "DÃ©pÃ´t": "Date",  # Encoded version of Dépôt
            "Livraison au plus tard": "ExpectedDeliveryDate"
        }
        
        df = df.rename(columns=column_mapping)
        
        # Check for required columns
        required_columns = ["ClientCode", "ProductCode", "Quantity", "Date"]
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            # Try alternative mapping
            alt_mapping = {}
            for col in df.columns:
                if "client" in col.lower() or "numéro" in col.lower():
                    alt_mapping[col] = "ClientCode"
                elif "code" in col.lower() and "article" in col.lower():
                    alt_mapping[col] = "ProductCode"
                elif "cadence" in col.lower():
                    alt_mapping[col] = "Quantity"
                elif "dépôt" in col.lower() or "depot" in col.lower():
                    alt_mapping[col] = "Date"
            
            df = df.rename(columns=alt_mapping)
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                available_cols = df.columns.tolist()
                return render_template_string(UPLOAD_FORM_HTML,
                                              error_message=f"Missing required columns: {missing_columns}. Available columns: {available_cols}")

        # Handle data conversion
        df = df.dropna(subset=required_columns)
        
        try:
            # Keep ClientCode and ProductCode as strings for alphanumeric values
            df["ClientCode"] = df["ClientCode"].astype(str).str.strip()
            df["ProductCode"] = df["ProductCode"].astype(str).str.strip()
            
            # Convert Date column to string and clean it
            df["Date"] = df["Date"].astype(str).str.strip()
            
            # Only convert Quantity to integer  
            df["Quantity"] = pd.to_numeric(df["Quantity"], errors='coerce').astype('Int64')
        except Exception as conv_error:
            return render_template_string(UPLOAD_FORM_HTML,
                                          error_message=f"Data conversion error: {conv_error}")
        
        # Only drop to these columns if ExpectedDeliveryDate is missing
        if "ExpectedDeliveryDate" in df.columns:
            df = df[["ClientCode", "ProductCode", "Date", "Quantity", "ExpectedDeliveryDate"]]
        else:
            df = df[["ClientCode", "ProductCode", "Date", "Quantity"]]

        preview_csv = temp_path.replace(".xlsx", ".csv")
        df.to_csv(preview_csv, index=False)

        table_html = df.head(20).to_html(index=False, classes="table")

        return render_template_string(UPLOAD_FORM_HTML,
                                      excel_table=table_html,
                                      transformed_table=None,
                                      preview_data=True,
                                      temp_file=os.path.basename(preview_csv),
                                      show_download=False,
                                      germany_error_message=None,
                                      germany_success_message=None,
                                      tunisia_success_message="Excel preview loaded successfully.",
                                      tunisia_error_message=None)
                                      
    except Exception as e:
        return render_template_string(UPLOAD_FORM_HTML,
                                      summary_table=None,
                                      show_download=False,
                                      germany_error_message=None,
                                      germany_success_message=None,
                                      tunisia_error_message=f"Excel parsing failed: {e}",
                                      tunisia_success_message=None)

    

@app.route('/insert_excel', methods=['POST'])
def insert_excel():
    temp_file = request.form.get("temp_file")
    week_number = request.form.get("week_number")
    file_path = os.path.join(OUTPUT_DIR, temp_file)

    try:
        if not week_number or not week_number.isdigit():
            raise ValueError("Week number is required and must be a valid number.")

        df = pd.read_csv(file_path)

        df["EDIWeekNumber"] = int(week_number)

        # Verify ExpectedDeliveryDate exists and has valid data
        if "ExpectedDeliveryDate" not in df.columns:
            raise ValueError("ExpectedDeliveryDate column missing from processed data. Please ensure your Excel file contains a 'Livraison au plus tard' column.")
        
        # Check for empty values
        empty_mask = df["ExpectedDeliveryDate"].isin(['', 'nan', 'None', 'NaT']) | df["ExpectedDeliveryDate"].isna()
        if empty_mask.any():
            print(f"Warning: {empty_mask.sum()} rows have empty ExpectedDeliveryDate values. Using Date as fallback.")
            df.loc[empty_mask, "ExpectedDeliveryDate"] = df.loc[empty_mask, "Date"]

        store_to_postgres_edi(df)

        return render_template_string(UPLOAD_FORM_HTML,
                                      excel_table=df.head(20).to_html(index=False, classes="table"),
                                      transformed_table=None,
                                      preview_data=False,
                                      show_download=False,
                                      germany_error_message=None,
                                      germany_success_message=None,
                                      tunisia_success_message=f"{len(df)} rows inserted with Week {week_number}.",
                                      tunisia_error_message=None)
    except Exception as e:
        return render_template_string(UPLOAD_FORM_HTML,
                                      summary_table=None,
                                      show_download=False,
                                      germany_error_message=None,
                                      germany_success_message=None,
                                      tunisia_error_message=f"Database insert failed: {e}",
                                      tunisia_success_message=None)


@app.route('/download/<filename>')
def download(filename):
    try:
        # Secure the filename to prevent directory traversal attacks
        safe_filename = secure_filename(filename)
        file_path = os.path.join(OUTPUT_DIR, safe_filename)
       
        # Normalize the path to handle any .. or similar attempts
        file_path = os.path.abspath(file_path)
        output_dir_abs = os.path.abspath(OUTPUT_DIR)
       
        # Ensure the file is within the OUTPUT_DIR
        if not file_path.startswith(output_dir_abs):
            logging.warning(f"Attempted access to file outside output directory: {filename}")
            return "Access denied", 403
       
        # Check if file exists
        if not os.path.exists(file_path):
            logging.error(f"File not found: {file_path}")
            logging.info(f"Available files in {OUTPUT_DIR}: {os.listdir(OUTPUT_DIR) if os.path.exists(OUTPUT_DIR) else 'Directory does not exist'}")
            return render_template_string(UPLOAD_FORM_HTML,
                                        summary_table=None,
                                        show_download=False,
                                        germany_error_message=f"File '{filename}' not found. Please try processing your file again.",
                                        germany_success_message=None,
                                        tunisia_error_message=None,
                                        tunisia_success_message=None)
       
        logging.info(f"Downloading file: {file_path}")
        return send_file(file_path, as_attachment=True, download_name=safe_filename)
       
    except Exception as e:
        logging.error(f"Error in download route: {e}")
        return render_template_string(UPLOAD_FORM_HTML,
                                    summary_table=None,
                                    show_download=False,
                                    germany_error_message=f"Error downloading file: {str(e)}",
                                    germany_success_message=None,
                                    tunisia_error_message=None,
                                    tunisia_success_message=None)
 


@app.route('/send_to_db', methods=['POST'])
def send_to_db():
    try:
        csv_file = request.form.get('csv_file')
        customer_name = request.form.get('customer_name')
        file_date = request.form.get('file_date')  # Get the date from form

        if not csv_file or not customer_name or not file_date:
            raise ValueError("CSV file, customer name, or file date is missing.")

        file_path = os.path.join(OUTPUT_DIR, secure_filename(csv_file))

        if not os.path.exists(file_path):
            raise FileNotFoundError(f"CSV file not found: {file_path}")

        df = pd.read_csv(file_path, sep=';')

        # Rename columns properly
        df = df.rename(columns={
            "Client name": "ClientName",
            "Client Material No": "ClientMaterialNo",
            "AVO Material No": "AVOMaterialNo",
            "Material Description": "MaterialDescription",
            "Quantity": "Quantity",
            "Date from": "Date",
            "Pickup date": "Date"
        })

        required_columns = ["ClientName", "ClientMaterialNo", "AVOMaterialNo", "MaterialDescription", "Quantity", "Date"]
        df = df[[col for col in required_columns if col in df.columns]]
        df.dropna(subset=required_columns, inplace=True)

        df["Quantity"] = pd.to_numeric(df["Quantity"], errors='coerce').fillna(0).astype(int)
        
        # Add the file date to all rows
        df["FileDate"] = file_date

        with engine.begin() as conn:
            for _, row in df.iterrows():
                stmt = text("""
                    INSERT INTO "EDIGermany" (
                        "ClientName", "ClientMaterialNo", "AVOMaterialNo",
                        "MaterialDescription", "Quantity", "Date", "FileDate"
                    )
                    VALUES (:ClientName, :ClientMaterialNo, :AVOMaterialNo,
                            :MaterialDescription, :Quantity, :Date, :FileDate)
                """)
                conn.execute(stmt, {
                    "ClientName": row["ClientName"],
                    "ClientMaterialNo": row["ClientMaterialNo"],
                    "AVOMaterialNo": row["AVOMaterialNo"],
                    "MaterialDescription": row["MaterialDescription"],
                    "Quantity": row["Quantity"],
                    "Date": row["Date"],
                    "FileDate": row["FileDate"]  # Add the file date
                })

        logging.info(f"{len(df)} rows inserted into EDIGermany table with file date {file_date}.")
        return render_template_string(UPLOAD_FORM_HTML,
                                      transformed_table=df.head(20).to_html(classes='table', index=False),
                                      show_download=False,
                                      germany_success_message=f"{len(df)} rows inserted into the database with file date {file_date}.",
                                      germany_error_message=None,
                                      tunisia_error_message=None,
                                      tunisia_success_message=None)

    except Exception as e:
        logging.error(f"Error sending data to database: {e}")
        return render_template_string(UPLOAD_FORM_HTML,
                                      transformed_table=None,
                                      show_download=False,
                                      germany_success_message=None,
                                      germany_error_message=f"Database insert failed: {str(e)}",
                                      tunisia_error_message=None,
                                      tunisia_success_message=None)


# Add a route to list available files (for debugging)
@app.route('/debug/files')
def list_files():
    if app.debug:
        try:
            files = os.listdir(OUTPUT_DIR)
            files_info = []
            for f in files:
                file_path = os.path.join(OUTPUT_DIR, f)
                size = os.path.getsize(file_path)
                mtime = datetime.fromtimestamp(os.path.getmtime(file_path))
                files_info.append(f"{f} - Size: {size} bytes - Modified: {mtime}")
            return f"Files in {OUTPUT_DIR}:<br>" + "<br>".join(files_info)
        except Exception as e:
            return f"Error listing files: {e}"
    else:
        return "Debug mode not enabled", 404
 
if __name__ == '__main__':
    # Ensure output directory exists on startup
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    logging.info(f"Output directory: {os.path.abspath(OUTPUT_DIR)}")
    app.run(debug=True, port=5000) 