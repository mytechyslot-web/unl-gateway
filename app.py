import os
import json
from flask import Flask, render_template, request
import gspread
from google.oauth2.service_account import Credentials

app = Flask(__name__)

# Connect to Google Sheets Ledger
def get_gspread_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    # We will set 'GCP_JSON' in the Cloud settings later
    creds_json = json.loads(os.environ.get("GCP_JSON"))
    creds = Credentials.from_service_account_info(creds_json, scopes=scopes)
    return gspread.authorize(creds)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/activate', methods=['POST'])
def activate():
    voucher_code = request.form.get('voucher')
    user_ip = request.remote_addr
    
    try:
        client = get_gspread_client()
        # Open your Lab_CRM sheet
        sheet = client.open("Lab_CRM").worksheet("Vouchers")
        cell = sheet.find(voucher_code)
        
        if cell:
            # Assuming Status is Column 3 (C)
            status = sheet.cell(cell.row, 3).value
            if status == "Available":
                sheet.update_cell(cell.row, 3, "USED")
                sheet.update_cell(cell.row, 4, user_ip)
                
                # Mark as ACTIVE in Master Sheet
                master = client.open("Lab_CRM").sheet1
                master.append_row([user_ip, "ACTIVE", "Voucher Used"])
                return "<h1>ACCESS GRANTED</h1><p>Neural Link Established.</p>"
        
        return "<h1>INVALID CODE</h1><p>Please contact the Lab Manager.</p>"
    except Exception as e:
        return f"<h1>SYSTEM ERROR</h1><p>{str(e)}</p>"

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=int(os.environ.get("PORT", 5000)))
