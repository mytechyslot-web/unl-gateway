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
                    # THE CRITICAL CHANGE: Update the sheet AND return SUCCESS immediately
                    sheet.update_cell(i + 1, 2, 'USED')
                    # This tells Flask to RENDER the success version of your template
                    return render_template('index.html', success=True, voucher=voucher_input)
        
        return "INVALID CODE. Please contact the Lab Manager."
    except Exception as e:
        # If there's a timeout or error, we see it here
        return f"SYSTEM ERROR: {str(e)}"
