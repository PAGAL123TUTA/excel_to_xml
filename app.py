from flask import Flask, render_template, request, send_file, redirect
import os
import pandas as pd
from datetime import datetime
import openpyxl
import warnings

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
ALLOWED_EXTENSIONS = {'xlsx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def convert_excel_to_xml(file_path, company_name):
    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip().str.upper()

    cash_keywords = ['cash']
    bank_keywords = ['bank', 'sbi', 'hdfc', 'icici', 'axis', 'kotak', 'pnb', 'punjab', 'and', 'sind']

    def is_cash_or_bank(name):
        if pd.isna(name): return False
        name = str(name).lower()
        return any(k in name for k in cash_keywords + bank_keywords)

    xml_header = f"""<ENVELOPE>
<HEADER>
  <TALLYREQUEST>Import Data</TALLYREQUEST>
</HEADER>
<BODY>
<IMPORTDATA>
<REQUESTDESC>
  <REPORTNAME>Vouchers</REPORTNAME>
  <STATICVARIABLES>
    <SVCURRENTCOMPANY>{company_name}</SVCURRENTCOMPANY>
  </STATICVARIABLES>
</REQUESTDESC>
<REQUESTDATA>
"""
    xml_footer = """</REQUESTDATA>
</IMPORTDATA>
</BODY>
</ENVELOPE>"""

    entries = ""
    for _, row in df.iterrows():
        if pd.isna(row['DATE']) or pd.isna(row['AMOUNT']):
            continue

        date = pd.to_datetime(row['DATE']).strftime('%Y%m%d')
        amount = float(row['AMOUNT'])
        narration = str(row['NARRATION']) if pd.notna(row['NARRATION']) else ""
        partyledger = str(row['PARTYLEDGERNAME']).strip()
        bankledger = str(row['BANKORCASHLEDGER']).strip()
        ref_number = str(row['REFERENCENUMBER']).strip() if pd.notna(row.get('REFERENCENUMBER')) else "On Account"

        vouchertype = str(row['VOUCHERTYPE']).strip().capitalize() if pd.notna(row.get('VOUCHERTYPE')) else ""
        if not vouchertype and pd.notna(row.get('TRANSACTIONTYPE')):
            txn_type = str(row['TRANSACTIONTYPE']).strip().upper()
            if txn_type == "CR":
                vouchertype = "Receipt"
            elif txn_type == "DR":
                vouchertype = "Payment"
        if is_cash_or_bank(partyledger) and is_cash_or_bank(bankledger):
            vouchertype = "Contra"
        if vouchertype not in ["Receipt", "Payment", "Contra"]:
            continue

        if vouchertype == "Receipt":
            entries_xml = [(partyledger, "No", amount), (bankledger, "Yes", -amount)]
        else:
            entries_xml = [(partyledger, "Yes", -amount), (bankledger, "No", amount)]

        ledger_entries = ""
        for ledgername, is_deemed_positive, amt in entries_xml:
            bill_type = "New Ref" if ref_number != "On Account" else "On Account"
            ledger_entries += f"""    <ALLLEDGERENTRIES.LIST>
        <LEDGERNAME>{ledgername}</LEDGERNAME>
        <ISDEEMEDPOSITIVE>{is_deemed_positive}</ISDEEMEDPOSITIVE>
        <AMOUNT>{amt:.2f}</AMOUNT>
        <BILLALLOCATIONS.LIST>
            <NAME>{ref_number}</NAME>
            <BILLTYPE>{bill_type}</BILLTYPE>
            <AMOUNT>{amt:.2f}</AMOUNT>
        </BILLALLOCATIONS.LIST>
    </ALLLEDGERENTRIES.LIST>"""

        voucher = f""" <TALLYMESSAGE xmlns:UDF="TallyUDF">
<VOUCHER VCHTYPE="{vouchertype}" ACTION="Create">
    <DATE>{date}</DATE>
    <VOUCHERTYPENAME>{vouchertype}</VOUCHERTYPENAME>
    <PARTYLEDGERNAME>{partyledger}</PARTYLEDGERNAME>
    <NARRATION>{narration}</NARRATION>
    {ledger_entries}
</VOUCHER>
</TALLYMESSAGE>"""
        entries += voucher

    return xml_header + entries + xml_footer

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return redirect(request.url)
    file = request.files['file']
    if file.filename == '':
        return redirect(request.url)
    if file and allowed_file(file.filename):
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(file_path)
        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
            ws = wb.active
            company_name = str(ws["I2"].value).strip()
            if not company_name:
                raise ValueError("Company name is empty")
        except Exception:
            return "Error: Company name not found or empty in I2", 400
        xml_data = convert_excel_to_xml(file_path, company_name)
        xml_filename = f"voucher_import_{company_name}.xml"
        with open(xml_filename, 'w', encoding='utf-8') as f:
            f.write(xml_data)
        return send_file(xml_filename, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
