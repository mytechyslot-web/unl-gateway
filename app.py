@app.route('/activate', methods=['POST'])
def activate():
    voucher_input = request.form.get('voucher', '').strip()
    try:
        sheet = client.open("Lab_CRM").worksheet("Vouchers")
        # Get all rows to bypass header-name dependency
        records = sheet.get_all_values()
        
        for i, row in enumerate(records):
            # row[0] is Column A, row[1] is Column B
            if len(row) >= 2:
                db_voucher = str(row[0]).strip()
                db_status = str(row[1]).strip()
                
                if db_voucher == voucher_input and db_status == 'ACTIVE':
                    # Success: Update the status to USED
                    sheet.update_cell(i + 1, 2, 'USED')
                    return render_template('index.html', success=True, voucher=voucher_input)
        
        return "INVALID CODE. Please contact the Lab Manager."
    except Exception as e:
        return f"SYSTEM ERROR: {str(e)}"
