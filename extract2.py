import os
import sys
import pdfplumber
import pandas as pd
import re
from collections import defaultdict
import logging

# ------------------ SUPPRESS PDFMINER WARNINGS ------------------
logging.getLogger("pdfminer").setLevel(logging.ERROR)

# ------------------ PDF DATA EXTRACTION ------------------
def extract_data_from_pdf_manual(pdf_path):
    rows = []
    with pdfplumber.open(pdf_path) as pdf:
        last_sku = ""
        last_sku_y = None
        system_no = None  # store system no once per PDF

        for page in pdf.pages:
            full_text = page.extract_text()

            # Extract System No, Order No, Tracking #
            if system_no is None:
                system_no_match = re.search(r"System No\.?\s*[: ]\s*([0-9]+)", full_text, re.I)
                system_no = system_no_match.group(1) if system_no_match else None

            order_no_match = re.search(r"Order No\.?\s*[: ]\s*([A-Za-z0-9]+)", full_text, re.I)
            tracking_no_match = re.search(r"Tracking #\s*[: ]\s*([A-Za-z0-9]+)", full_text, re.I)
            order_no = order_no_match.group(1) if order_no_match else None
            tracking_no = tracking_no_match.group(1) if tracking_no_match else None

            words = page.extract_words()
            lines = defaultdict(list)
            for w in words:
                top_key = round(w['top'] / 2) * 2
                lines[top_key].append(w)

            sorted_lines = sorted(lines.items(), key=lambda x: x[0])

            for _, line_words in sorted_lines:
                line_words = sorted(line_words, key=lambda w: w['x0'])
                sku = ""
                description = ""
                qty = ""

                for w in line_words:
                    x = w['x0']
                    y = w['top']
                    text = w['text'].strip()

                    # SKU pattern
                    if 10 <= x <= 40 and re.match(r"^[0-9A-Za-z]{3,5}$", text):
                        sku = text
                        last_sku = sku
                        last_sku_y = y
                    elif 40 < x < 220:
                        description += text
                    elif 220 <= x <= 260:
                        if text.isdigit():
                            qty = text

                # Use last SKU if blank but close vertically
                #if not sku and last_sku_y is not None:
                    #if any(abs(w['top'] - last_sku_y) <= 30 for w in line_words):
                        #sku = last_sku

                if sku and qty:
                    rows.append({
                        "System No": system_no,
                        "Order No": order_no,
                        "Tracking #": tracking_no,
                        "SKU": sku,
                        "Description": description.strip(),
                        "QTY ORD": qty,
                        # "Source File": os.path.basename(pdf_path)
                    })

    return rows

# ------------------ PROCESS FOLDER ------------------
def process_folder_manual(pdf_folder_path, output_folder_path, output_filename="output_manual.xlsx", limit=None):
    all_rows = []

    all_files = os.listdir(pdf_folder_path)

    # Sort PDFs by NEWEST FIRST (modification time)
    pdf_files = sorted(
        [f for f in all_files if f.lower().endswith(".pdf")],
        key=lambda f: os.path.getmtime(os.path.join(pdf_folder_path, f)),
        reverse=True
    )

    if limit is not None:
        pdf_files = pdf_files[:limit]

    if not pdf_files:
        print("No PDF files found.")
        return

    for file in pdf_files:
        pdf_path = os.path.join(pdf_folder_path, file)
        print(f"Processing: {pdf_path}")
        rows = extract_data_from_pdf_manual(pdf_path)
        all_rows.extend(rows)

    df = pd.DataFrame(all_rows)
    df.drop_duplicates(inplace=True)
    # df['SKU'] = pd.to_numeric(df['SKU'], errors='coerce').fillna(0).astype(int)
    df['QTY ORD'] = pd.to_numeric(df['QTY ORD'], errors='coerce').fillna(0).astype(int)

    os.makedirs(output_folder_path, exist_ok=True)
    output_excel_path = os.path.join(output_folder_path, output_filename)
    df.to_excel(output_excel_path, index=False)

    print(f"\n✔ DONE! Excel saved as: {output_excel_path}")
    return df

# ------------------ MAIN ------------------
if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python extract2.py <PDF_FOLDER_PATH> <OUTPUT_FOLDER_PATH> [LIMIT]")
        sys.exit(1)

    pdf_folder_path = sys.argv[1]
    output_folder_path = sys.argv[2]
    limit = int(sys.argv[3]) if len(sys.argv) > 3 and sys.argv[3].isdigit() else None

    process_folder_manual(pdf_folder_path, output_folder_path, limit=limit)
