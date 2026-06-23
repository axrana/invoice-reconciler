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
    """Extract last numeric value from a text line"""
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


def get_next_non_empty_line(lines, start_index, max_lookahead=5):
    """Get next useful line below a heading"""
    for j in range(start_index + 1, min(start_index + 1 + max_lookahead, len(lines))):
        candidate = lines[j].strip()
        if candidate:
            return candidate
    return None


def extract_pdf_data(pdf_path):
    data = {
        "BillNo": None,
        "Date": None,
        "PDF_Buyer_Name": None,
        "PDF_Consignee_Name": None,
        "PDF_NetAmt_Taxable": None,
        "PDF_NetAmt_TotalRow": None,
        "PDF_Amount_Invoice": None,
        "PDF_Amount_Terms": None,
        "TDS": None,
        "GST": None,
        "Days": None
    }

    with pdfplumber.open(pdf_path) as pdf:
        all_text = ""
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                all_text += text + "\n"

    lines = [line.strip() for line in all_text.split("\n") if line.strip()]

    for i, line in enumerate(lines):

        # 1. BillNo: 9 digits after LS/
        if "Invoice No." in line:
            bill_match = re.search(r"LS/(\d{9})", line, re.IGNORECASE)
            if bill_match:
                data["BillNo"] = bill_match.group(1)

        # 2. Date
        if "Invoice Date" in line:
            date_match = re.search(r"(\d{2}/\d{2}/\d{4})", line)
            if date_match:
                data["Date"] = date_match.group(1)

        # 3. Buyer Name - same next-line logic as consignee
        if "Details of Buyer" in line and "Billed to" in line:
            data["PDF_Buyer_Name"] = get_next_non_empty_line(lines, i)

        # 4. Consignee Name
        if "Details of Consignee" in line and "Shipped to" in line:
            data["PDF_Consignee_Name"] = get_next_non_empty_line(lines, i)

        # 5. NetAmt location 1
        if "Total Taxable Amt in INR" in line and "1.00" in line:
            data["PDF_NetAmt_Taxable"] = extract_numbers_from_text(line)

        # 6. NetAmt location 2
        if "Total" in line:
            val = extract_numbers_from_text(line)
            if val is not None:
                data["PDF_NetAmt_TotalRow"] = val

        # 7. Amount location 1
        if "Total Invoice Amount" in line and "(INR)" in line:
            data["PDF_Amount_Invoice"] = extract_numbers_from_text(line)

        # 8. Amount location 2 + Days
        if "Terms of Delivery and Payment from the date of invoice" in line:
            below_line = get_next_non_empty_line(lines, i)
            if below_line:
                days_match = re.search(r"^\d{1,3}", below_line)
                if days_match:
                    data["Days"] = int(days_match.group(0))

                if "INR" in below_line:
                    parts = below_line.split("INR")
                    if len(parts) > 1:
                        data["PDF_Amount_Terms"] = extract_numbers_from_text(parts[-1])

        # 9. TDS
        if "LESS 0.10%" in line and "TDS" in line:
            data["TDS"] = extract_numbers_from_text(line)

        # 10. GST
        if "Total Tax Amount" in line and "(INR)" in line:
            data["GST"] = extract_numbers_from_text(line)

    return data


def run_reconciliation():
    print("🚀 Running Offline Verification Engine...")

    if not os.path.exists(EXCEL_FILE):
        print(f"❌ Error: Cannot find your Excel sheet named '{EXCEL_FILE}'")
        return

    excel_df = pd.read_excel(EXCEL_FILE)
    excel_df.columns = excel_df.columns.str.strip()
    excel_df['BillNo'] = excel_df['BillNo'].astype(str).str.strip()

    # KEEP THIS EXACTLY AS REQUESTED
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
    pdf_df.columns = pdf_df.columns.str.strip()
    pdf_df['BillNo'] = pdf_df['BillNo'].astype(str).str.strip()

    print("\n🔍 Evaluating variances...")
    merged = pd.merge(excel_df, pdf_df, on="BillNo", suffixes=('_Excel', '_PDF'), how='outer')

    results = []

    for _, row in merged.iterrows():
        status = "Match"
        notes = []

        pdf_file = row['PDF_File'] if 'PDF_File' in merged.columns else None
        epname = row['EPName'] if 'EPName' in merged.columns else None

        if pd.isna(pdf_file):
            status = "Missing PDF"
            notes.append("No matching physical invoice found.")
        elif pd.isna(epname):
            status = "Missing Excel Entry"
            notes.append("Invoice found in PDF files but row missing from master ledger.")
        else:
            excel_name = str(row['EPName']).strip().lower()
            pdf_buyer = str(row['PDF_Buyer_Name']).strip().lower() if not pd.isna(row.get('PDF_Buyer_Name')) else ""
            pdf_consignee = str(row['PDF_Consignee_Name']).strip().lower() if not pd.isna(row.get('PDF_Consignee_Name')) else ""

            # Party name should match either buyer or consignee
            if excel_name not in pdf_buyer and excel_name not in pdf_consignee:
                status = "Discrepancy"
                notes.append("Party Name mismatch")

            excel_net = float(row['NetAmt']) if not pd.isna(row.get('NetAmt')) else 0.0
            excel_amt = float(row['Amount']) if not pd.isna(row.get('Amount')) else 0.0
            excel_tds = float(row['TDS_Excel']) if 'TDS_Excel' in merged.columns and not pd.isna(row.get('TDS_Excel')) else (
                float(row['TDS']) if 'TDS' in merged.columns and not pd.isna(row.get('TDS')) else 0.0
            )
            excel_gst = float(row['GST_Excel']) if 'GST_Excel' in merged.columns and not pd.isna(row.get('GST_Excel')) else (
                float(row['GST']) if 'GST' in merged.columns and not pd.isna(row.get('GST')) else 0.0
            )

            pdf_net_taxable = float(row['PDF_NetAmt_Taxable']) if not pd.isna(row.get('PDF_NetAmt_Taxable')) else None
            pdf_net_totalrow = float(row['PDF_NetAmt_TotalRow']) if not pd.isna(row.get('PDF_NetAmt_TotalRow')) else None
            pdf_amt_invoice = float(row['PDF_Amount_Invoice']) if not pd.isna(row.get('PDF_Amount_Invoice')) else None
            pdf_amt_terms = float(row['PDF_Amount_Terms']) if not pd.isna(row.get('PDF_Amount_Terms')) else None
            pdf_tds = float(row['TDS']) if 'TDS' in merged.columns and not pd.isna(row.get('TDS')) else None
            pdf_gst = float(row['GST']) if 'GST' in merged.columns and not pd.isna(row.get('GST')) else None

            # NetAmt: only 2 checks
            net_match_found = False
            if pdf_net_taxable is not None and abs(excel_net - pdf_net_taxable) <= 5.0:
                net_match_found = True
            if pdf_net_totalrow is not None and abs(excel_net - pdf_net_totalrow) <= 5.0:
                net_match_found = True
            if not net_match_found:
                status = "Discrepancy"
                notes.append(
                    f"NetAmt mismatch (Excel: {excel_net:.2f}, Taxable: {pdf_net_taxable if pdf_net_taxable is not None else 'NA'}, TotalRow: {pdf_net_totalrow if pdf_net_totalrow is not None else 'NA'})"
                )

            # Amount: only 2 checks
            amt_match_found = False
            if pdf_amt_invoice is not None and abs(excel_amt - pdf_amt_invoice) <= 5.0:
                amt_match_found = True
            if pdf_amt_terms is not None and abs(excel_amt - pdf_amt_terms) <= 5.0:
                amt_match_found = True
            if not amt_match_found:
                status = "Discrepancy"
                notes.append(
                    f"Amount mismatch (Excel: {excel_amt:.2f}, Invoice: {pdf_amt_invoice if pdf_amt_invoice is not None else 'NA'}, Terms: {pdf_amt_terms if pdf_amt_terms is not None else 'NA'})"
                )

            # TDS
            if pdf_tds is not None:
                if abs(excel_tds - pdf_tds) > 5.0:
                    status = "Discrepancy"
                    notes.append(f"TDS mismatch (Excel: {excel_tds:.2f}, PDF: {pdf_tds:.2f})")

            # GST
            if pdf_gst is not None:
                if abs(excel_gst - pdf_gst) > 5.0:
                    status = "Discrepancy"
                    notes.append(f"GST mismatch (Excel: {excel_gst:.2f}, PDF: {pdf_gst:.2f})")

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
