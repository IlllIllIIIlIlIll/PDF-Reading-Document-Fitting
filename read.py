import os
import sys
import re
import pdfplumber
import pandas as pd

# Define file paths
root_dir = os.path.dirname(__file__)  # Get script's directory
excel_file = os.path.join(root_dir, "temp", "CEK LIST PEB.xlsx")
pdf_folder = os.path.join(root_dir, "temp", "ALL2022")
ocr_folder = os.path.join(root_dir, "temp", "OCR")  # New OCR folder

# Ensure OCR directory exists
os.makedirs(ocr_folder, exist_ok=True)

# Read Excel file from the "ALL" sheet
df = pd.read_excel(excel_file, sheet_name="ALL", dtype=str)  # Ensure all data is read as strings

# Check if "INV NO" column exists
if "INV NO" not in df.columns:
    raise ValueError("Column 'INV NO' not found in the 'ALL' sheet of the Excel file.")

# Use "INV NO" directly as the document name
df["DOC NAME"] = df["INV NO"].astype(str)  # Ensure it's a string

# List all PDF files in the ALL2022 folder
all_pdf_files = os.listdir(pdf_folder)

# Prepare new DataFrame columns
peb_numbers = []
inv_numbers = []
file_names = []
is_splitted_from = []
ok_flags = []

# Regex pattern to extract "VAM-XXXXX"
vam_pattern = re.compile(r"VAM-(\d+)")

# Process each document
for doc_name in df["DOC NAME"]:
    matching_files = [f for f in all_pdf_files if doc_name.lower() in f.lower()]

    # Check if it's a split document
    doc_range = None
    range_start, range_end = None, None

    if "-" in doc_name:
        range_match = vam_pattern.findall(doc_name)
        if len(range_match) >= 2:
            range_start, range_end = int(range_match[0]), int(range_match[1])
            doc_range = doc_name

    if not matching_files:
        # print(f"File not found: {doc_name}")
        peb_numbers.append(None)
        inv_numbers.append(doc_name)
        file_names.append(None)
        is_splitted_from.append(None)
        ok_flags.append("NO")
    else:
        pdf_path = os.path.join(pdf_folder, matching_files[0])
        file_names.append(matching_files[0])

        print(f"\n🔍 Processing: {pdf_path}\n")

        # Generate OCR output path inside "OCR" folder
        ocr_pdf_path = os.path.join(ocr_folder, matching_files[0])

        # Check if the PDF already contains text
        with pdfplumber.open(pdf_path) as pdf:
            page_text = pdf.pages[0].extract_text()

        if page_text:
            print("✅ PDF already contains text. Skipping OCR.")
            text_source = pdf_path  # Use the original file
        else:
            # Apply OCR only if needed
            if not os.path.exists(ocr_pdf_path):
                print(f"⚠️ No OCR file found, running OCR on {pdf_path}...")
                os.system(f'python -m ocrmypdf --skip-text "{pdf_path}" "{ocr_pdf_path}"')

            # Use OCR-processed PDF for text extraction
            text_source = ocr_pdf_path

        # Try reading text from the selected PDF (either original or OCR-processed)
        with pdfplumber.open(text_source) as pdf:
            page_text = pdf.pages[0].extract_text()

        # Print extracted OCR result
        if page_text:
            print("\n📝 OCR Extracted Text:")
            print(page_text[:1000])  # Print first 1000 characters to keep output manageable
            print("\n" + "-"*80 + "\n")
        else:
            print(f"⚠️ No readable text found in {text_source}.")

        if not page_text:
            print(f"⚠️ No readable text found in {text_source}.")
            peb_numbers.append(None)
            inv_numbers.append(doc_name)
            is_splitted_from.append(None)
            ok_flags.append("NO")
            continue

        # Check if the document contains "BC 3.0 PEMBERITAHUAN EKSPOR BARANG"
        if "BC 3.0 PEMBERITAHUAN EKSPOR BARANG" not in page_text:
            print(f"⚠️ Document does not contain expected header: {text_source}")
            peb_numbers.append(None)
            inv_numbers.append(doc_name)
            is_splitted_from.append(None)
            ok_flags.append("NO")
            continue

        # Extract PEB number (Nomor Pendaftaran)
        peb_number = None
        for line in page_text.split("\n"):
            if "Nomor Pendaftaran" in line:
                peb_number = line.split(":")[-1].strip()
                break

        # Check if the doc_name matches "22. Nomor & Tgl Invoice"
        doc_found = False
        for line in page_text.split("\n"):
            if "22. Nomor & Tgl Invoice" in line and doc_name in line:
                doc_found = True
                break

        # Handling split documents
        if doc_range:
            doc_found = False
            for line in page_text.split("\n"):
                if "22. Nomor & Tgl Invoice" in line:
                    for num in range(range_start, range_end + 1):
                        if f"VAM-{num}" in line:
                            doc_found = True
                            break

        # Set "isSplittedFrom" field
        if doc_range and doc_found:
            is_splitted_from.append(f"PEB {doc_range}.pdf")
        else:
            is_splitted_from.append(None)

        # Set "OK?" column
        ok_flags.append("OK" if doc_found else "NO")
        peb_numbers.append(peb_number)
        inv_numbers.append(doc_name)

# Create output DataFrame
output_df = pd.DataFrame({
    "PEB": peb_numbers,
    "INV": inv_numbers,
    "FILENAME (IF EXIST)": file_names,
    "isSplittedFrom": is_splitted_from,
    "OK?": ok_flags
})

# Save the updated DataFrame as a new sheet in the Excel file
output_file = os.path.join(root_dir, "temp", "Updated_CEK_LIST_PEB.xlsx")

with pd.ExcelWriter(output_file, mode="w", engine="xlsxwriter") as writer:
    output_df.to_excel(writer, sheet_name="Checked Data", index=False)

print(f"\n✅ Processing complete. Results saved to {output_file}.")
