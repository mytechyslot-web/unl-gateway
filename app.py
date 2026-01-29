import os
import json
from flask import Flask, render_template, request
import gspread
from google.oauth2.service_account import Credentials

app = Flask(__name__)

# Google Sheets Configuration
# This pulls the JSON credentials you pasted into Render's Environment Variables
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
gcp_json = os.environ.get('GCP_JSON')
creds_dict = json.loads(gcp_json)
creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
client = gspread.authorize(creds)

@app.route('/')
def index():
    # Renders the beautiful glassmorphism landing page
    return render_template('index.html')

@app.route('/activate', methods=['POST'])
def activate():
    # .strip() and .upper() handle accidental spaces or lowercase typing
    voucher_input = request.form.get('voucher', '').strip().upper()
    
    try:
        # Access the sheet and tab
        sheet = client.open("Lab_CRM").worksheet("Vouchers")
        records = sheet.get_all_values()
        
        for i, row in enumerate(records):
            # Column A is row[0], Column B is row[1]
            if len(row) >= 2:
                db_voucher = str(row[0]).strip().upper()
                db_status = str(row[1]).strip().upper()
                
                # Check for match
                if db_voucher == voucher_input and db_status == 'ACTIVE':
                    # Success: Flip status to USED in Google Sheets
                    sheet.update_cell(i + 1, 2, 'USED')
                    return render_template('index.html', success=True, voucher=voucher_input)
        
        # If loop finishes without finding the code
        return "INVALID CODE. Please contact the Lab Manager."
        
    except Exception as e:
        # Diagnostic report if the link breaks
        return f"SYSTEM ERROR: {str(e)}"

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000)
