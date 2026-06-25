import os
import re
import glob
import pdfplumber
import pandas as pd

# --- CONFIGURATION ---
EXCEL_FILE = "master_list.xlsx"
INVOICE_FOLDER = "Invoices"
OUTPUT_FILE = "reconciliation_report.xlsx"

# Tolerance in rupees for NetAmt, Amount, TDS, GST (Excel already *1000)
AMOUNT_TOL = 3.0


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


def get_next_non_empty_line(lines, start_index, max_lookahead=8):
    """Return the next non-empty line after start_index."""
    for j in range(start_index + 1, min(start_index + 1 + max_lookahead, len(lines))):
        candidate = lines[j].strip()
        if candidate:
            return candidate
    return None


def normalize_name(name):
    """Normalize party name for exact comparison: strip, collapse spaces, lowercase."""
    if not name or pd.isna(name):
        return ""
    name = " ".join(str(name).strip().split())
    return name.lower()


def normalize_billno(value):
    """Normalize BillNo like LS/112633489 -> 112633489 for comparison."""
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

        # --- BillNo: Invoice No. : LS/112633489 -> 112633489 ---
        if "INVOICE NO." in upper_line:
            bill_match = re.search(r"LS/\s*(\d{8,9})", line, re.IGNORECASE)
            if bill_match:
                data["BillNo"] = bill_match.group(1)

        # --- Date: Invoice Date : DD/MM/YYYY ---
        if "INVOICE DATE" in upper_line:
            date_match = re.search(r"(\d{2}/\d{2}/\d{4})", line)
            if date_match:
                data["Date"] = date_match.group(1)

        # --- Buyer Name: find first valid line BELOW header ---
        if line.strip() == "Details of Buyer ( Billed to)":
            j = i + 1
            while j < len(lines):
                candidate = lines[j].strip()
                upper_cand = candidate.upper()

                # stop if we hit start of consignee section
                if "DETAILS OF" in upper_cand and "CONSIGNEE" in upper_cand:
                    break

                # skip empty or non-name lines
                if candidate and not (
                    "INVOICE NO." in upper_cand
                    or "INVOICE DATE" in upper_cand
                    or "GSTIN/UID" in upper_cand
                    or "FINANCIAL YEAR" in upper_cand
                ):
                    data["PDF_Buyer_Name"] = candidate
                    break
                j += 1

                               # --- Consignee Name: match on "Consignee ( Shipped to)" only ---
        if "CONSIGNEE ( SHIPPED TO)" in upper_line:
            # look for the next non-empty line
            j = i + 1
            while j < len(lines):
                candidate = lines[j].strip()
                if candidate:  # non-empty
                    upper_cand = candidate.upper()
                    # if the very next non-empty line is already GSTIN, then no consignee name line
                    if "GSTIN/UID" in upper_cand:
                        break
                    # otherwise use this line as consignee name
                    data["PDF_Consignee_Name"] = candidate
                    break
                j += 1

        # --- NetAmt candidate 1: Total Taxable Amt in INR @ 1.00 ... ---
        if "TOTAL TAXABLE AMT IN INR" in upper_line and "1.00" in upper_line:
            parts = re.split(r"Total Taxable Amt in INR\s*@\s*1\.00", line, flags=re.IGNORECASE)
            if len(parts) > 1:
                data["PDF_NetAmt_Taxable"] = extract_numbers_from_text(parts[1])
            else:
                data["PDF_NetAmt_Taxable"] = extract_numbers_from_text(line)

        # --- NetAmt candidate 2: C&F total row (Total ... C&F ...) ---
        if "TOTAL" in upper_line and "C&F" in upper_line:
            data["PDF_NetAmt_TotalRow"] = extract_numbers_from_text(line)

        # --- Amount candidate 1: Total Invoice Amount (INR) ---
        if "TOTAL INVOICE AMOUNT" in upper_line and "(INR)" in upper_line:
            parts = re.split(r"Total Invoice Amount\s*\(INR\)", line, flags=re.IGNORECASE)
            if len(parts) > 1:
                data["PDF_Amount_Invoice"] = extract_numbers_from_text(parts[1])
            else:
                data["PDF_Amount_Invoice"] = extract_numbers_from_text(line)

        # --- Terms: exact header, use LAST amount in next non-empty line ---
        if line.strip().startswith("Terms of Delivery and Payment from the date of invoice"):
            below = get_next_non_empty_line(lines, i, max_lookahead=8)
            if below:
                # Optional Days at start (e.g. "5 DAYS ...")
                days_match = re.search(r"^\s*(\d{1,3})", below)
                if days_match:
                    data["Days"] = int(days_match.group(1))
                # Terms amount: last number in the line (e.g. DELIVERY ON ADVANCE INR 77,597.00)
                terms_amt = extract_numbers_from_text(below)
                if terms_amt is not None:
                    data["PDF_Amount_Terms"] = terms_amt

        # --- TDS ---
        if "LESS 0.10%" in upper_line and "TDS" in upper_line:
            data["TDS"] = extract_numbers_from_text(line)

        # --- GST (Total Tax Amount) ---
        if "TOTAL TAX AMOUNT" in upper_line and "(INR)" in upper_line:
            data["GST"] = extract_numbers_from_text(line)

    return data


def run_reconciliation():
    print("🚀 Running Offline Verification Engine...")

    if not os.path.exists(EXCEL_FILE):
        print(f"❌ Error: Cannot find your Excel sheet named '{EXCEL_FILE}'")
        return

    # --- Load Excel ---
    excel_df = pd.read_excel(EXCEL_FILE)
    excel_df.columns = excel_df.columns.str.strip()

    # Normalize BillNo like LS/112633489 -> 112633489
    excel_df["BillNo"] = excel_df["BillNo"].apply(normalize_billno)

    # Scale Excel values x1000 (as required)
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
            # --- Exact party name check (first line, normalized) ---
            excel_name_norm = normalize_name(row["EPName"])
            buyer_name_norm = normalize_name(row.get("PDF_Buyer_Name"))
            consignee_name_norm = normalize_name(row.get("PDF_Consignee_Name"))

            if excel_name_norm:
                if not (
                    excel_name_norm == buyer_name_norm
                    or excel_name_norm == consignee_name_norm
                ):
                    status = "Discrepancy"
                    notes.append(
                        "Party Name mismatch (Excel EPName not exactly equal to Buyer/Consignee name)."
                    )

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

            if net1 is not None or net2 is not None:
                pdf_pair_consistent = True
                if net1 is not None and net2 is not None:
                    pdf_pair_consistent = abs(net1 - net2) <= AMOUNT_TOL

                m1 = net1 is not None and abs(excel_net - net1) <= AMOUNT_TOL
                m2 = net2 is not None and abs(excel_net - net2) <= AMOUNT_TOL

                net_ok = pdf_pair_consistent and (m1 or m2)

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

            if amt1 is not None or amt2 is not None:
                pdf_pair_consistent = True
                if amt1 is not None and amt2 is not None:
                    pdf_pair_consistent = abs(amt1 - amt2) <= AMOUNT_TOL

                m1 = amt1 is not None and abs(excel_amt - amt1) <= AMOUNT_TOL
                m2 = amt2 is not None and abs(excel_amt - amt2) <= AMOUNT_TOL

                amt_ok = pdf_pair_consistent and (m1 or m2)

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

            # --- GST ---
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
            excel_days_val = row.get("Days_Excel", row.get("Days"))
            if not pd.isna(excel_days_val) and pdf_days is not None and not pd.isna(pdf_days):
                try:
                    excel_days_int = int(excel_days_val)
                    pdf_days_int = int(pdf_days)
                    if excel_days_int != pdf_days_int:
                        status = "Discrepancy"
                        notes.append(
                            f"Days mismatch (Excel: {excel_days_int}, PDF: {pdf_days_int})"
                        )
                except Exception:
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
