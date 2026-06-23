import os
import re
import glob
import pdfplumber
import pandas as pd

# --- CONFIGURATION (Adjust file name if needed) ---
EXCEL_FILE = "master_list.xlsx"  # Change this if your file name is different
INVOICE_FOLDER = "Invoices"
OUTPUT_FILE = "reconciliation_report.xlsx"

def extract_pdf_data(pdf_path):
    """Extracts required matching fields from a single offline PDF invoice"""
    data = {
        "BillNo": None, "Date": None, "EPName": None, 
        "Amount": None, "TDS": None, "GST": None, "NetAmt": None
    }
    
    with pdfplumber.open(pdf_path) as pdf:
        all_text = ""
        all_tables = []
        
        # Scrape text and tables across all pages
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                all_text += text + "\n"
            tables = page.extract_tables()
            if tables:
                all_tables.extend(tables)

        # 1. BillNo: Find "Invoice No. :  LS/" followed by a 9-digit number
        bill_match = re.search(r"Invoice\s*No\.\s*:\s*LS/(\d{9})", all_text, re.IGNORECASE)
        if bill_match:
            data["BillNo"] = f"LS/{bill_match.group(1)}"

        # 2. Date: Find "Invoice Date : DD/MM/YYYY"
        date_match = re.search(r"Invoice\s*Date\s*:\s*(\d{2}/\d{2}/\d{4})", all_text, re.IGNORECASE)
        if date_match:
            data["Date"] = date_match.group(1)

        # 3. EPName: Look below Consignee or Buyer details
        lines = all_text.split("\n")
        for i, line in enumerate(lines):
            if "Details of Consignee" in line or "Details of Buyer" in line:
                # Grab the next non-empty line as the Company Name
                for next_line in lines[i+1:i+5]:
                    if next_line.strip():
                        data["EPName"] = next_line.strip()
                        break
                break

        # 4. TDS: After "LESS 0.10% TDS IN INR"
        tds_match = re.search(r"LESS\s*0\.10%\s*TDS\s*IN\s*INR\s*[:\-]?\s*([\d,.]+)", all_text, re.IGNORECASE)
        if tds_match:
            data["TDS"] = float(tds_match.group(1).replace(",", ""))

        # 5. GST: After "Total Tax Amount (INR)"
        gst_match = re.search(r"Total\s*Tax\s*Amount\s*\(INR\)\s*[:\-]?\s*([\d,.]+)", all_text, re.IGNORECASE)
        if gst_match:
            data["GST"] = float(gst_match.group(1).replace(",", ""))

        # 6. NetAmt: After "Total Invoice Amount (INR)"
        net_match = re.search(r"Total\s*Invoice\s*Amount\s*\(INR\)\s*[:\-]?\s*([\d,.]+)", all_text, re.IGNORECASE)
        if net_match:
            data["NetAmt"] = float(net_match.group(1).replace(",", ""))

        # 7. Amount: Total row, last column (Check tables)
        for table in all_tables:
            for row in table:
                if row and any(cell and "Total" in str(cell) for cell in row):
                    # Filter out empty cells in the row and grab the last one
                    valid_cells = [cell for cell in row if cell is not None and str(cell).strip()]
                    if valid_cells:
                        last_cell = valid_cells[-1].replace(",", "")
                        try:
                            data["Amount"] = float(re.findall(r"[\d.]+", last_cell)[0])
                        except IndexError:
                            pass
                    break

    return data

def run_reconciliation():
    print("🚀 Starting Offline PDF Extraction...")
    
    # Check if files exist
    if not os.path.exists(EXCEL_FILE):
        print(f"❌ Error: Could not find '{EXCEL_FILE}' in this directory.")
        return
        
    # Read Excel Sheets
    excel_df = pd.read_excel(EXCEL_FILE)
    # Ensure data types match for clean comparison
    excel_df['BillNo'] = excel_df['BillNo'].astype(str).str.strip()
    
    # Scan PDFs
    pdf_files = glob.glob(os.path.join(INVOICE_FOLDER, "*.pdf"))
    if not pdf_files:
        print(f"❌ Error: No PDFs found inside the '{INVOICE_FOLDER}' folder.")
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

    print("\n🔍 Cross-Referencing data against Excel...")
    
    # Merge datasets based on BillNo
    merged = pd.merge(excel_df, pdf_df, on="BillNo", suffixes=('_Excel', '_PDF'), how='outer')

    # Status Matching Logic
    results = []
    for idx, row in merged.iterrows():
        status = "Match"
        notes = []

        if pd.isna(row['PDF_File']):
            status = "Missing PDF"
            notes.append("Invoice exists in Excel but no matching PDF found.")
        elif pd.isna(row['Date_Excel']):
            status = "Missing Excel Record"
            notes.append("PDF exists but Invoice is missing in Excel list.")
        else:
            # 1. Check Amount within +-5 tolerance
            amt_excel = float(row['Amount_Excel']) if not pd.isna(row['Amount_Excel']) else 0
            amt_pdf = float(row['Amount_PDF']) if not pd.isna(row['Amount_PDF']) else 0
            if abs(amt_excel - amt_pdf) > 5:
                status = "Discrepancy"
                notes.append(f"Amount mismatch (Excel: {amt_excel}, PDF: {amt_pdf})")

            # 2. Check NetAmt, GST, TDS mismatches
            for col in ['NetAmt', 'GST', 'TDS']:
                ex_val = float(row[f'{col}_Excel']) if not pd.isna(row[f'{col}_Excel']) else 0
                pdf_val = float(row[f'{col}_PDF']) if not pd.isna(row[f'{col}_PDF']) else 0
                if abs(ex_val - pdf_val) > 0.5:  # small cents/paise tolerance
                    status = "Discrepancy"
                    notes.append(f"{col} mismatch")

        row['Reconciliation_Status'] = status
        row['Discrepancy_Notes'] = ", ".join(notes) if notes else "All values match"
        results.append(row)

    output_df = pd.DataFrame(results)
    
    # Clean up columns order for output
    final_cols = ['BillNo', 'Reconciliation_Status', 'Discrepancy_Notes', 'PDF_File'] + \
                 [c for c in output_df.columns if c not in ['BillNo', 'Reconciliation_Status', 'Discrepancy_Notes', 'PDF_File']]
    output_df = output_df[final_cols]

    output_df.to_excel(OUTPUT_FILE, index=False)
    print(f"🎉 Success! Report saved cleanly as: {OUTPUT_FILE}")

if __name__ == "__main__":
    run_reconciliation()
