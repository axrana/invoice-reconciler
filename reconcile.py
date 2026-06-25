import os
import re
import glob
import pdfplumber
import pandas as pd

# --- CONFIGURATION ---
EXCEL_FILE = "master_list.xlsx"
INVOICE_FOLDER = "Invoices"
OUTPUT_FILE = "reconciliation_report.xlsx"

# Amount tolerance in rupees (remember: Excel already *1000)
AMOUNT_TOL = 3.0  # NetAmt, Amount, TDS, GST


def extract_numbers_from_text(text):
    """Extract last numeric value from a text line."""
    if not text:
        return None
    found = re.findall(r"[\d,.]+", text)
    if not found:
        return None
    clean_num = found[-1].replace(",", "")
    try:
        return float(clean_num)
    except ValueError:
        return None


def get_next_non_empty_lines(lines, start_index, max_lines=2, stop_markers=None, max_lookahead=10):
    """
    Collect up to `max_lines` non-empty lines below a header,
    stopping early if we hit any stop marker (e.g. GSTIN/UID, Invoice No.).
    """
    if stop_markers is None:
        stop_markers = []

    collected = []
    for j in range(start_index + 1, min(start_index + 1 + max_lookahead, len(lines))):
        candidate = lines[j].strip()
        if not candidate:
            continue

        upper_candidate = candidate.upper()
        if any(marker.upper() in upper_candidate for marker in stop_markers):
            break

        collected.append(candidate)
        if len(collected) >= max_lines:
            break

    return " ".join(collected).strip() if collected else None


def normalize_billno(value):
    """Normalize BillNo like LS/99999999 -> 99999999 for comparison."""
    if pd.isna(value):
        return ""
    value = str(value).strip()
    match = re.search(r"(\d{8,9})", value)
    return match.group(1) if match else value


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
        "Days": None,
    }

    with pdfplumber.open(pdf_path) as pdf:
        all_text = ""
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                all_text += text + "\n"

    lines = [line.strip() for line in all_text.split("\n") if line.strip()]

    for i, line in enumerate(lines):
        upper_line = line.upper()

        # --- BillNo: Invoice No. : LS/99999999 -> 9 digits ---
        if "INVOICE NO." in upper_line:
            bill_match = re.search(r"LS/\s*(\d{8,9})", line, re.IGNORECASE)
            if bill_match:
                data["BillNo"] = bill_match.group(1)

        # --- Date ---
        if "INVOICE DATE" in upper_line:
            date_match = re.search(r"(\d{2}/\d{2}/\d{4})", line)
            if date_match:
                data["Date"] = date_match.group(1)

        # --- Buyer Name: up to 2 lines, stop at GSTIN/UID or Invoice No. ---
        if "DETAILS OF BUYER" in upper_line and "BILLED TO" in upper_line:
            data["PDF_Buyer_Name"] = get_next_non_empty_lines(
                lines,
                i,
                max_lines=2,
                stop_markers=["GSTIN/UID", "INVOICE NO."]
            )

        # --- Consignee Name: up to 2 lines, stop at GSTIN/UID ---
        if "DETAILS OF  CONSIGNEE" in upper_line and "SHIPPED TO" in upper_line:
            data["PDF_Consignee_Name"] = get_next_non_empty_lines(
                lines,
                i,
                max_lines=2,
                stop_markers=["GSTIN/UID"]
            )

        # --- NetAmt candidate 1: Total Taxable Amt in INR @ 1.00 ... ---
        if "TOTAL TAXABLE AMT IN INR" in upper_line and "1.00" in upper_line:
            # Use text after the label if possible
            parts = re.split(r"Total Taxable Amt in INR\s*@\s*1\.00", line, flags=re.IGNORECASE)
            if len(parts) > 1:
                data["PDF_NetAmt_Taxable"] = extract_numbers_from_text(parts[1])
            else:
                data["PDF_NetAmt_Taxable"] = extract_numbers_from_text(line)

        # --- NetAmt candidate 2: C&F / Total row ---
        if "TOTAL" in upper_line and "C&F" in upper_line:
            data["PDF_NetAmt_TotalRow"] = extract_numbers_from_text(line)

        # --- Amount candidate 1: Total Invoice Amount (INR) ---
        if "TOTAL INVOICE AMOUNT" in upper_line and "(INR)" in upper_line:
            parts = re.split(r"Total Invoice Amount\s*\(INR\)", line, flags=re.IGNORECASE)
            if len(parts) > 1:
                data["PDF_Amount_Invoice"] = extract_numbers_from_text(parts[1])
            else:
                data["PDF_Amount_Invoice"] = extract_numbers_from_text(line)

        # --- Terms block: Days + Amount candidate 2 ("5 DAYS FIX INR ...") ---
        if "TERMS OF DELIVERY AND PAYMENT FROM THE DATE OF INVOICE" in upper_line:
            # Next non-empty line holds Days and FIX INR amount
            below = get_next_non_empty_lines(lines, i, max_lines=1)
            if below:
                # Days: first 1-3 digits at start
                days_match = re.search(r"^\s*(\d{1,3})", below)
                if days_match:
                    data["Days"] = int(days_match.group(1))

                # Amount after "INR"
                if "INR" in below.upper():
                    parts = re.split(r"INR", below, flags=re.IGNORECASE)
                    if len(parts) > 1:
                        data["PDF_Amount_Terms"] = extract_numbers_from_text(parts[1])

        # --- TDS ---
        if "LESS 0.10%" in upper_line and "TDS" in upper_line:
            data["TDS"] = extract_numbers_from_text(line)

        # --- GST (Total Tax Amount) ---
        if "TOTAL TAX AMOUNT" in upper_line and "(INR)" in upper_line:
            data["GST"] = extract_numbers_from_text(line)

    return data


def run_reconciliation():
    print("🚀 Running Offline Verification Engine with double-location checks...")

    if not os.path.exists(EXCEL_FILE):
        print(f"❌ Error: Cannot find your Excel sheet named '{EXCEL_FILE}'")
        return

    # --- Load Excel ---
    excel_df = pd.read_excel(EXCEL_FILE)
    excel_df.columns = excel_df.columns.str.strip()

    # Normalize BillNo like LS/99999999 -> 99999999
    excel_df["BillNo"] = excel_df["BillNo"].apply(normalize_billno)

    # Scale Excel values x1000 (as you require)
    excel_cols_to_scale = ["NetAmt", "Amount", "TDS", "GST"]
    for col in excel_cols_to_scale:
        if col in excel_df.columns:
            excel_df[col] = (
                pd.to_numeric(excel_df[col], errors="coerce")
                .fillna(0.0)
                * 1000.0
            )

    # --- Scan PDFs ---
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
    if not pdf_df.empty:
        pdf_df.columns = pdf_df.columns.str.strip()
        pdf_df["BillNo"] = pdf_df["BillNo"].apply(normalize_billno)

    print("\n🔍 Evaluating variances...")
    merged = pd.merge(
        excel_df,
        pdf_df,
        on="BillNo",
        suffixes=("_Excel", "_PDF"),
        how="outer",
    )

    results = []

    for _, row in merged.iterrows():
        status = "Match"
        notes = []

        pdf_file = row.get("PDF_File")
        epname = row.get("EPName")

        # --- Presence checks ---
        if pd.isna(pdf_file):
            status = "Missing PDF"
            notes.append("No matching PDF found for this BillNo.")
        elif pd.isna(epname):
            status = "Missing Excel Entry"
            notes.append("PDF exists but no corresponding row in Excel.")
        else:
            # --- Party name check (Buyer + Consignee, 2 lines each) ---
            excel_name = str(row["EPName"]).strip().lower()

            pdf_buyer = str(row.get("PDF_Buyer_Name", "")).strip().lower()
            pdf_consignee = str(row.get("PDF_Consignee_Name", "")).strip().lower()

            if excel_name and excel_name not in pdf_buyer and excel_name not in pdf_consignee:
                status = "Discrepancy"
                notes.append("Party Name mismatch (Excel EPName not found in Buyer/Consignee block).")

            # --- Core numeric fields from Excel (already x1000) ---
            excel_net = float(row["NetAmt"]) if not pd.isna(row.get("NetAmt")) else 0.0
            excel_amt = float(row["Amount"]) if not pd.isna(row.get("Amount")) else 0.0
            excel_tds = float(row["TDS"]) if not pd.isna(row.get("TDS")) else 0.0
            excel_gst = float(row["GST"]) if not pd.isna(row.get("GST")) else 0.0

            # --- NetAmt candidates ---
            net1 = row.get("PDF_NetAmt_Taxable")
            net2 = row.get("PDF_NetAmt_TotalRow")
            net1 = float(net1) if not pd.isna(net1) else None
            net2 = float(net2) if not pd.isna(net2) else None

            net_ok = False
            if net1 is not None or net2 is not None:
                # Ensure the two PDF values themselves are consistent if both exist
                pdf_pair_consistent = True
                if net1 is not None and net2 is not None:
                    pdf_pair_consistent = abs(net1 - net2) <= AMOUNT_TOL

                # Check Excel vs each candidate
                m1 = net1 is not None and abs(excel_net - net1) <= AMOUNT_TOL
                m2 = net2 is not None and abs(excel_net - net2) <= AMOUNT_TOL

                if pdf_pair_consistent and (m1 or m2):
                    net_ok = True

                if not net_ok:
                    status = "Discrepancy"
                    notes.append(
                        f"NetAmt mismatch (Excel: {excel_net:.2f}, "
                        f"Taxable: {net1 if net1 is not None else 'NA'}, "
                        f"TotalRow: {net2 if net2 is not None else 'NA'})"
                    )

            # --- Amount (gross) candidates ---
            amt1 = row.get("PDF_Amount_Invoice")
            amt2 = row.get("PDF_Amount_Terms")
            amt1 = float(amt1) if not pd.isna(amt1) else None
            amt2 = float(amt2) if not pd.isna(amt2) else None

            amt_ok = False
            if amt1 is not None or amt2 is not None:
                pdf_pair_consistent = True
                if amt1 is not None and amt2 is not None:
                    pdf_pair_consistent = abs(amt1 - amt2) <= AMOUNT_TOL

                m1 = amt1 is not None and abs(excel_amt - amt1) <= AMOUNT_TOL
                m2 = amt2 is not None and abs(excel_amt - amt2) <= AMOUNT_TOL

                if pdf_pair_consistent and (m1 or m2):
                    amt_ok = True

                if not amt_ok:
                    status = "Discrepancy"
                    notes.append(
                        f"Amount mismatch (Excel: {excel_amt:.2f}, "
                        f"Invoice: {amt1 if amt1 is not None else 'NA'}, "
                        f"Terms: {amt2 if amt2 is not None else 'NA'})"
                    )

            # --- TDS ---
            tds_pdf = row.get("TDS")
            tds_pdf = float(tds_pdf) if not pd.isna(tds_pdf) else None
            if tds_pdf is not None:
                if abs(excel_tds - tds_pdf) > AMOUNT_TOL:
                    status = "Discrepancy"
                    notes.append(
                        f"TDS mismatch (Excel: {excel_tds:.2f}, PDF: {tds_pdf:.2f})"
                    )

            # --- GST (Total Tax Amount) ---
            gst_pdf = row.get("GST")
            gst_pdf = float(gst_pdf) if not pd.isna(gst_pdf) else None
            if gst_pdf is not None:
                if abs(excel_gst - gst_pdf) > AMOUNT_TOL:
                    status = "Discrepancy"
                    notes.append(
                        f"GST mismatch (Excel: {excel_gst:.2f}, PDF: {gst_pdf:.2f})"
                    )

            # --- Days: exact match, no tolerance ---
            pdf_days = row.get("Days")
            if not pd.isna(row.get("Days")) and not pd.isna(row.get("Days_Excel", row.get("Days"))):
                # Excel Days may be in 'Days' column (no suffix) or 'Days_Excel' after merge
                excel_days_val = row.get("Days_Excel", row.get("Days"))
                if not pd.isna(excel_days_val) and pdf_days is not None:
                    try:
                        excel_days_int = int(excel_days_val)
                        pdf_days_int = int(pdf_days)
                        if excel_days_int != pdf_days_int:
                            status = "Discrepancy"
                            notes.append(
                                f"Days mismatch (Excel: {excel_days_int}, PDF: {pdf_days_int})"
                            )
                    except Exception:
                        # If conversion fails, skip Days check
                        pass

        row["Reconciliation_Status"] = status
        row["Discrepancy_Notes"] = ", ".join(notes) if notes else "All data points verified"
        results.append(row)

    output_df = pd.DataFrame(results)

    priority_cols = ["BillNo", "Reconciliation_Status", "Discrepancy_Notes", "PDF_File"]
    final_order = priority_cols + [c for c in output_df.columns if c not in priority_cols]
    output_df = output_df[final_order]

    output_df.to_excel(OUTPUT_FILE, index=False)
    print(f"🎉 Complete! View details in: {OUTPUT_FILE}")


if __name__ == "__main__":
    run_reconciliation()
