import os
import re
import glob
import pdfplumber
import pandas as pd
import pytesseract
from PIL import Image
from pdf2image import convert_from_path

EXCEL_FILE     = "master_list_purchase.xlsx"
INVOICE_FOLDER = "Invoices_Purchase"
OUTPUT_FILE    = "purchase_reconciliation_report.xlsx"
AMOUNT_TOL     = 3.0

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def extract_numbers_from_text(text):
    if not text:
        return None
    found = re.findall(r"[\d,.]+", text)
    if not found:
        return None
    clean = found[-1].replace(",", "")
    try:
        return float(clean)
    except ValueError:
        return None

def normalize_billno(value):
    if pd.isna(value):
        return ""
    value = str(value).strip()
    match = re.search(r"(\d{6,12})", value)
    return match.group(1) if match else value

def file_to_text(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".pdf":
        raw = ""
        try:
            with pdfplumber.open(file_path) as pdf:
                raw = "\n".join(p.extract_text() or "" for p in pdf.pages)
        except Exception:
            raw = ""
        if len(raw.strip()) > 80:
            return raw
        try:
            images = convert_from_path(file_path, dpi=300)
            return "\n".join(pytesseract.image_to_string(img) for img in images)
        except Exception as e:
            return f"OCR_ERROR: {e}"
    elif ext in (".jpg", ".jpeg", ".png", ".tiff", ".bmp", ".webp"):
        try:
            img = Image.open(file_path)
            return pytesseract.image_to_string(img)
        except Exception as e:
            return f"OCR_ERROR: {e}"
    return ""

FIELD_PATTERNS = {
    "BillNo": [
        r"(?:invoice\s*no[.:]?|bill\s*no[.:]?|invoice\s*#|ref\s*no[.:]?)\s*[:/]?\s*([A-Z0-9/\-]+)",
    ],
    "Date": [
        r"(?:invoice\s*date|bill\s*date|date)[\s:]+(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2,4})",
        r"(\d{2}/\d{2}/\d{4})",
    ],
    "VendorName": [
        r"(?:from|seller|vendor|supplier|billed\s*by)[\s:\n]+([A-Za-z][^\n]{3,60})",
        r"(?:m/s\.?|m/s)[\s]+([^\n]{3,60})",
    ],
    "TaxableAmt": [
        r"(?:total\s*taxable|taxable\s*amount|net\s*amount|sub\s*total)[\s:]+([\\d,\.]+)",
    ],
    "TotalAmt": [
        r"(?:grand\s*total|total\s*amount|invoice\s*total|net\s*payable|amount\s*payable|total\s*invoice)[\s:]+(?:inr|rs\.?)?\s*([\d,\.]+)",
    ],
    "GST": [
        r"(?:total\s*gst|total\s*tax|igst|cgst\s*\+\s*sgst|tax\s*amount)[\s:]+(?:inr|rs\.?)?\s*([\d,\.]+)",
    ],
    "TDS": [
        r"(?:tds|tax\s*deducted)[\s:@\d\.%]+([\d,\.]+)",
    ],
}

def extract_fields(text):
    data = {k: None for k in FIELD_PATTERNS}
    data["RawText"] = text[:500]
    text_lower = text.lower()
    for field, patterns in FIELD_PATTERNS.items():
        for pattern in patterns:
            match = re.search(pattern, text_lower, re.IGNORECASE | re.MULTILINE)
            if match:
                raw_val = match.group(1).strip()
                if field in ("TaxableAmt", "TotalAmt", "GST", "TDS"):
                    data[field] = extract_numbers_from_text(raw_val)
                else:
                    data[field] = raw_val
                break
    if data["BillNo"]:
        data["BillNo"] = normalize_billno(data["BillNo"])
    return data

def run_reconciliation():
    print("\n🚀 Axrana Purchase Invoice Reconciler (OCR Edition)\n")

    if not os.path.exists(EXCEL_FILE):
        print(f"❌ Cannot find '{EXCEL_FILE}'. Place it next to this script.")
        return

    excel_df = pd.read_excel(EXCEL_FILE)
    excel_df.columns = excel_df.columns.str.strip()

    if "BillNo" not in excel_df.columns:
        print("❌ Excel must have a BillNo column.")
        return

    excel_df["BillNo"] = excel_df["BillNo"].apply(normalize_billno)

    for col in ["NetAmt", "Amount", "TDS", "GST"]:
        if col in excel_df.columns:
            excel_df[col] = pd.to_numeric(excel_df[col], errors="coerce").fillna(0.0) * 1000.0

    file_patterns = [
        "*.pdf", "*.PDF", "*.jpg", "*.JPG", "*.jpeg", "*.JPEG",
        "*.png", "*.PNG", "*.tiff", "*.TIFF", "*.bmp", "*.BMP"
    ]
    invoice_files = []
    for pat in file_patterns:
        invoice_files += glob.glob(os.path.join(INVOICE_FOLDER, pat))

    if not invoice_files:
        print(f"❌ No invoice files found in '{INVOICE_FOLDER}' folder.")
        return

    extracted = []
    for fpath in invoice_files:
        fname = os.path.basename(fpath)
        fext  = os.path.splitext(fname)[1].upper()
        print(f"  📄 [{fext}] {fname} ...", end=" ")
        try:
            text   = file_to_text(fpath)
            fields = extract_fields(text)
            fields["InvoiceFile"] = fname
            fields["FileType"]    = fext
            extracted.append(fields)
            print(f"✅  BillNo={fields['BillNo']}  Total={fields['TotalAmt']}")
        except Exception as e:
            print(f"❌ Error: {e}")
            extracted.append({"InvoiceFile": fname, "FileType": fext, "BillNo": None})

    inv_df = pd.DataFrame(extracted)
    if not inv_df.empty:
        inv_df["BillNo"] = inv_df["BillNo"].apply(
            lambda x: normalize_billno(x) if x else None
        )

    print("\n🔍 Reconciling against master list...")
    merged = pd.merge(excel_df, inv_df, on="BillNo", how="outer")

    results = []
    for _, row in merged.iterrows():
        status = "Match"
        notes  = []

        inv_file = row.get("InvoiceFile")
        bill_no  = row.get("BillNo")

        if pd.isna(inv_file) or inv_file is None:
            status = "Missing Invoice File"
            notes.append("No invoice file found for this BillNo in Excel.")
        elif pd.isna(bill_no) or bill_no is None or bill_no == "":
            status = "BillNo Not Extracted"
            notes.append("Invoice file present but could not extract BillNo.")
        else:
            def check_amt(label, excel_col, pdf_col):
                e_val = row.get(excel_col)
                p_val = row.get(pdf_col)
                if pd.isna(e_val) or pd.isna(p_val):
                    return
                e_val, p_val = float(e_val), float(p_val)
                if abs(e_val - p_val) > AMOUNT_TOL:
                    notes.append(
                        f"{label} mismatch (Excel: {e_val:.2f}, Invoice: {p_val:.2f})"
                    )

            check_amt("Total Amount",   "Amount", "TotalAmt")
            check_amt("Taxable Amount", "NetAmt",  "TaxableAmt")
            check_amt("GST",            "GST",     "GST")
            check_amt("TDS",            "TDS",     "TDS")

            if notes:
                status = "Discrepancy"

        row["Reconciliation_Status"] = status
        row["Notes"] = ", ".join(notes) if notes else "All verified ✅"
        results.append(row)

    out_df = pd.DataFrame(results)
    priority = ["BillNo", "Reconciliation_Status", "Notes", "InvoiceFile", "FileType"]
    final_cols = priority + [c for c in out_df.columns if c not in priority]
    out_df = out_df[[c for c in final_cols if c in out_df.columns]]

    out_df.to_excel(OUTPUT_FILE, index=False)
    print(f"\n🎉 Done! Open '{OUTPUT_FILE}' to see your report.")

if __name__ == "__main__":
    run_reconciliation()
