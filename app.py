import os
import json
from flask import Flask, render_template, request
import gspread
from google.oauth2.service_account import Credentials

app = Flask(__name__)

# System Authentication
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
gcp_json = os.environ.get('GCP_JSON')
creds_dict = json.loads(gcp_json)
creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
client = gspread.authorize(creds)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/activate', methods=['POST'])
def activate():
    voucher_input = request.form.get('voucher', '').strip().upper()
    try:
        sheet = client.open("Lab_CRM").worksheet("Vouchers")
        records = sheet.get_all_values()
        
        for i, row in enumerate(records):
            if len(row) >= 2:
                db_voucher = str(row[0]).strip().upper()
                db_status = str(row[1]).strip().upper()
                
                if db_voucher == voucher_input and db_status == 'ACTIVE':
                    # STEP 1: Update the Ledger
                    sheet.update_cell(i + 1, 2, 'USED')
                    # STEP 2: Tell the HTML to show the Success Page
                    return render_template('index.html', success=True, voucher=voucher_input)
        
        return render_template('index.html', error="INVALID CODE. PLEASE CONTACT LAB MANAGER.")
    except Exception as e:
        return f"SYSTEM ERROR: {str(e)}"

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000)
