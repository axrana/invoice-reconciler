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
        return 0.0
    # Find numbers that look like integers or decimals (e.g., 1,234.50 or 500)
    found = re.findall(r"[\d,.]+", text)
    if found:
        # Use the last numeric element found in that snippet
        clean_num = found[-1].replace(",", "")
        try:
            return float(clean_num)
        except ValueError:
            return 0.0
    return 0.0

def extract_pdf_data(pdf_path):
    data = {
        "BillNo": None, "Date": None, "EPName": None, 
        "NetAmt": None, "Amount": None, "TDS": None, "GST": None, "Days": None
    }
    
    with pdfplumber.open(pdf_path) as pdf:
        all_text = ""
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                all_text += text + "\n"

        lines = [line.strip() for line in all_text.split("\n") if line.strip()]

        for i, line in enumerate(lines):
            
            # 1. BillNo: Find "Invoice No. :" and pull only the 9 digits after "LS/"
            if "Invoice No." in line:
                bill_match = re.search(r"LS/(\d{9})", line, re.IGNORECASE)
                if bill_match:
                    data["BillNo"] = bill_match.group(1)

            # 2. Date: Find "Invoice Date :" and parse DD/MM/YYYY
            if "Invoice Date" in line:
                date_match = re.search(r"(\d{2}/\d{2}/\d{4})", line)
                if date_match:
                    data["Date"] = date_match.group(1)

            # 3. EPName: Target line directly below "Details of Buyer ( Billed to)"
            if "Details of Buyer" in line and "Billed to" in line:
                if i + 1 < len(lines):
                    data["EPName"] = lines[i + 1]

            # 4. NetAmt (Priority 1): Check "Total Taxable Amt in INR @ 1.00"
            if "Total Taxable Amt in INR" in line and "1.00" in line:
                data["NetAmt"] = extract_numbers_from_text(line) * 1000

            # 5. NetAmt (Priority 2 / Fallback): Look for "Total" row and take the final/last amount segment
            if "Total" in line and data["NetAmt"] is None:
                data["NetAmt"] = extract_numbers_from_text(line) * 1000

            # 6. Amount (Priority 1): Check "Total Invoice Amount (INR)"
            if "Total Invoice Amount" in line and "(INR)" in line:
                data["Amount"] = extract_numbers_from_text(line) * 1000

            # 7. Amount & Days (Fallback via Terms block)
            if "Terms of Delivery and Payment from the date of invoice" in line:
                if i + 1 < len(lines):
                    below_line = lines[i + 1]
                    
                    # Days: Pull first sequence of digits up to 3 characters long
                    days_match = re.search(r"^\d{1,3}", below_line)
                    if days_match:
                        data["Days"] = int(days_match.group(0))
                    
                    # Amount Fallback: If not found yet, get the last numbers showing after "INR"
                    if data["Amount"] is None and "INR" in below_line:
                        parts = below_line.split("INR")
                        if len(parts) > 1:
                            data["Amount"] = extract_numbers_from_text(parts[-1]) * 1000

            # 8. TDS: Extract after "LESS 0.10% TDS IN INR"
            if "LESS 0.10%" in line and "TDS" in line:
                data["TDS"] = extract_numbers_from_text(line) * 1000

            # 9. GST: Extract after "Total Tax Amount (INR)"
            if "Total Tax Amount" in line and "(INR)" in line:
                data["GST"] = extract_numbers_from_text(line) * 1000

    return data

def run_reconciliation():
    print("🚀 Running Offline Extraction Engine...")
    
    if not os.path.exists(EXCEL_FILE):
        print(f"❌ Error: Cannot find your Excel sheet named '{EXCEL_FILE}'")
        return
        
    excel_df = pd.read_excel(EXCEL_FILE)
    
    # Standardize Excel BillNo formats to clean text string for exact matching
    excel_df['BillNo'] = excel_df['BillNo'].astype(str).str.strip()
    
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

    print("\n🔍 Evaluating variances...")
    merged = pd.merge(excel_df, pdf_df, on="BillNo", suffixes=('_Excel', '_PDF'), how='outer')

    results = []
    for idx, row in merged.iterrows():
        status = "Match"
        notes = []

        if pd.isna(row['PDF_File']):
            status = "Missing PDF"
            notes.append("No matching physical invoice found.")
        elif pd.isna(row['EPName_Excel']):
            status = "Missing Excel Entry"
            notes.append("Invoice found in PDF files but row missing from master ledger.")
        else:
            # Check numeric differences with a tolerance allowance of +/- 5 rupees
            for numeric_col in ['NetAmt', 'Amount', 'TDS', 'GST']:
                val_ex = float(row[f'{numeric_col}_Excel']) if not pd.isna(row[f'{numeric_col}_Excel']) else 0.0
                val_pdf = float(row[f'{numeric_col}_PDF']) if not pd.isna(row[f'{numeric_col}_PDF']) else 0.0
                if abs(val_ex - val_pdf) > 5.0:
                    status = "Discrepancy"
                    notes.append(f"{numeric_col} mismatch (Excel: {val_ex:.2f}, PDF: {val_pdf:.2f})")

        row['Reconciliation_Status'] = status
        row['Discrepancy_Notes'] = ", ".join(notes) if notes else "All data points verified"
        results.append(row)

    output_df = pd.DataFrame(results)
    
    # Order final file output cleanly
    priority_cols = ['BillNo', 'Reconciliation_Status', 'Discrepancy_Notes', 'PDF_File']
    final_order = priority_cols + [col for col in output_df.columns if col not in priority_cols]
    output_df = output_df[final_order]

    output_df.to_excel(OUTPUT_FILE, index=False)
    print(f"🎉 Complete! View details in: {OUTPUT_FILE}")

if __name__ == "__main__":
    run_reconciliation()
