import os
import re
import glob
import pdfplumber
import pandas as pd

# --- CONFIGURATION ---
EXCEL_FILE = "master_list.xlsx" 
INVOICE_FOLDER = "Invoices"
OUTPUT_FILE = "reconciliation_report.xlsx"

def extract_numbers_from_text(text):
    """Helper to cleanly extract a numeric amount out of a mixed text string"""
    if not text:
        return None
    found = re.findall(r"[\d,.]+", text)
    if found:
        clean_num = found[-1].replace(",", "")
        try:
            return float(clean_num)
        except ValueError:
            return None
    return None

def extract_pdf_data(pdf_path):
    data = {
        "BillNo": None, "Date": None, "PDF_Buyer_Name": None, "PDF_Consignee_Name": None,
        "PDF_NetAmt_Taxable": None, "PDF_NetAmt_TotalRow": None,
        "PDF_Amount_Invoice": None, "PDF_Amount_Terms": None,
        "TDS": None, "GST": None, "Days": None
    }
    
    with pdfplumber.open(pdf_path) as pdf:
        all_text = ""
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                all_text += text + "\n"

        lines = [line.strip() for line in all_text.split("\n") if line.strip()]

        for i, line in enumerate(lines):
            
            # 1. BillNo: 9 digits after "LS/"
            if "Invoice No." in line:
                bill_match = re.search(r"LS/(\d{9})", line, re.IGNORECASE)
                if bill_match:
                    data["BillNo"] = bill_match.group(1)

            # 2. Date: DD/MM/YYYY
            if "Invoice Date" in line:
                date_match = re.search(r"(\d{2}/\d{2}/\d{4})", line)
                if date_match:
                    data["Date"] = date_match.group(1)

            # 3. EPName Verification (Dual Fields Extraction)
            # Buyer Name
            if "Details of Buyer" in line and "Billed to" in line:
                if i + 1 < len(lines):
                    data["PDF_Buyer_Name"] = lines[i + 1].strip()
            
            # Consignee Name
            if "Details of Consignee" in line and "Shipped to" in line:
                if i + 1 < len(lines):
                    data["PDF_Consignee_Name"] = lines[i + 1].strip()

            # 4. NetAmt Place 1: "Total Taxable Amt in INR @ 1.00"
            if "Total Taxable Amt in INR" in line and "1.00" in line:
                data["PDF_NetAmt_Taxable"] = extract_numbers_from_text(line)

            # 5. NetAmt Place 2: "Total" row last column
            if "Total" in line:
                val = extract_numbers_from_text(line)
                if val is not None:
                    data["PDF_NetAmt_TotalRow"] = val

            # 6. Amount Place 1: "Total Invoice Amount (INR)"
            if "Total Invoice Amount" in line and "(INR)" in line:
                data["PDF_Amount_Invoice"] = extract_numbers_from_text(line)

            # 7. Amount Place 2 & Days: Terms block
            if "Terms of Delivery and Payment from the date of invoice" in line:
                if i + 1 < len(lines):
                    below_line = lines[i + 1]
                    
                    days_match = re.search(r"^\d{1,3}", below_line)
                    if days_match:
                        data["Days"] = int(days_match.group(0))
                    
                    if "INR" in below_line:
                        parts = below_line.split("INR")
                        if len(parts) > 1:
                            data["PDF_Amount_Terms"] = extract_numbers_from_text(parts[-1])

            # 8. TDS
            if "LESS 0.10%" in line and "TDS" in line:
                data["TDS"] = extract_numbers_from_text(line)

            # 9. GST
            if "Total Tax Amount" in line and "(INR)" in line:
                data["GST"] = extract_numbers_from_text(line)

    return data

def run_reconciliation():
    print("🚀 Running Offline Verification Engine (Dual Party Name Checks Enabled)...")
    
    if not os.path.exists(EXCEL_FILE):
        print(f"❌ Error: Cannot find your Excel sheet named '{EXCEL_FILE}'")
        return
        
    excel_df = pd.read_excel(EXCEL_FILE)
    excel_df['BillNo'] = excel_df['BillNo'].astype(str).str.strip()
    
    # Scale Excel values x1000 for comparison
    excel_cols_to_scale = ['NetAmt', 'Amount', 'TDS', 'GST']
    for col in excel_cols_to_scale:
        if col in excel_df.columns:
            excel_df[col] = pd.to_numeric(excel_df[col], errors='coerce').fillna(0.0) * 1000
    
    pdf_files = glob.glob(os.path.join(INVOICE_FOLDER, "*.pdf"))
    if not pdf_files:
        print(f"❌ Error: Place your invoices into the '{INVOICE_FOLDER}' directory first.")
        return

    extracted_records = []
    for pdf_path in pdf_files:
        filename = os.path.basename(pdf_path)
        print(f"📄 Processing: {filename}")
        pdf_data = extract_pdf_data(pdf_path)
        pdf_data["PDF_File"] = filename
        extracted_records.append(pdf_data)

    pdf_df = pd.DataFrame(extracted_records)
    pdf_df['BillNo'] = pdf_df['BillNo'].astype(str).str.strip()

    print("\n🔍 Evaluating variances across all multiple layout locations...")
    merged = pd.merge(excel_df, pdf_df, on="BillNo", suffixes=('_Excel', '_PDF'), how='outer')

    results = []
    for idx, row in merged.iterrows():
        status = "Match"
        notes = []

        if pd.isna(row['PDF_File']):
            status = "Missing PDF"
            notes.append("No matching physical invoice found.")
        elif pd.isna(row['EPName']):
            status = "Missing Excel Entry"
            notes.append("Invoice found in PDF files but row missing from master ledger.")
        else:
            excel_name = str(row['EPName']).strip().lower()
            pdf_buyer = str(row['PDF_Buyer_Name']).strip().lower() if not pd.isna(row['PDF_Buyer_Name']) else ""
            pdf_consignee = str(row['PDF_Consignee_Name']).strip().lower() if not pd.isna(row['PDF_Consignee_Name']) else ""
            
            # --- VERIFY PARTY NAMES (Should match either Buyer OR Consignee) ---
            if excel_name not in pdf_buyer and excel_name not in pdf_consignee:
                status = "Discrepancy"
                notes.append("Party Name mismatch (Excel name does not match Buyer or Consignee lines)")

            excel_net = float(row['NetAmt']) if not pd.isna(row['NetAmt']) else 0.0
            excel_amt = float(row['Amount']) if not pd.isna(row['Amount']) else 0.0
            
            # --- DUAL LOCATION CHECK FOR NETAMT ---
            pdf_net_taxable = float(row['PDF_NetAmt_Taxable']) if not pd.isna(row['PDF_NetAmt_Taxable']) else None
            pdf_net_totalrow = float(row['PDF_NetAmt_TotalRow']) if not pd.isna(row['PDF_NetAmt_TotalRow']) else None
            
            if pdf_net_taxable is not None and abs(excel_net - pdf_net_taxable) > 5.0:
                status = "Discrepancy"
                notes.append(f"NetAmt Mismatch at Taxable Line (Excel: {excel_net:.2f}, PDF: {pdf_net_taxable:.2f})")
            if pdf_net_totalrow is not None and abs(excel_net - pdf_net_totalrow) > 5.0:
                status = "Discrepancy"
                notes.append(f"NetAmt Mismatch at Total Row (Excel: {excel_net:.2f}, PDF: {pdf_net_totalrow:.2f})")

            # --- DUAL LOCATION CHECK FOR AMOUNT ---
            pdf_amt_invoice = float(row['PDF_Amount_Invoice']) if not pd.isna(row['PDF_Amount_Invoice']) else None
            pdf_amt_terms = float(row['PDF_Amount_Terms']) if not pd.isna(row['PDF_Amount_Terms']) else None
            
            if pdf_amt_invoice is not None and abs(excel_amt - pdf_amt_invoice) > 5.0:
                status = "Discrepancy"
                notes.append(f"Amount Mismatch at Invoice Row (Excel: {excel_amt:.2f}, PDF: {pdf_amt_invoice:.2f})")
            if pdf_amt_terms is not None and abs(excel_amt - pdf_amt_terms) > 5.0:
                status = "Discrepancy"
                notes.append(f"Amount Mismatch at Terms Row (Excel: {excel_amt:.2f}, PDF: {pdf_amt_terms:.2f})")

            # --- STANDARD TDS & GST CHECKS ---
            for numeric_col in ['TDS', 'GST']:
                val_ex = float(row[numeric_col]) if not pd.isna(row[numeric_col]) else 0.0
                val_pdf = float(row[f'{numeric_col}_PDF']) if not pd.isna(row[f'{numeric_col}_PDF']) else 0.0
                if abs(val_ex - val_pdf) > 5.0:
                    status = "Discrepancy"
                    notes.append(f"{numeric_col} mismatch (Excel: {val_ex:.2f}, PDF: {val_pdf:.2f})")

        row['Reconciliation_Status'] = status
        row['Discrepancy_Notes'] = ", ".join(notes) if notes else "All data points verified"
        results.append(row)

    output_df = pd.DataFrame(results)
    
    priority_cols = ['BillNo', 'Reconciliation_Status', 'Discrepancy_Notes', 'PDF_File']
    final_order = priority_cols + [col for col in output_df.columns if col not in priority_cols]
    output_df = output_df[final_order]

    output_df.to_excel(OUTPUT_FILE, index=False)
    print(f"🎉 Complete! View details in: {OUTPUT_FILE}")

if __name__ == "__main__":
    run_reconciliation()
