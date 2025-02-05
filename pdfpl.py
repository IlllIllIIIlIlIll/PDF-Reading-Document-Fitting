import os
import sys
import re
import pdfplumber
import pandas as pd

root_dir = os.path.dirname(__file__) 
excel_file = os.path.join(root_dir, "temp", "CEK LIST PEB.xlsx")
pdf_folder = os.path.join(root_dir, "temp", "ALL2022")
ocr_folder = os.path.join(root_dir, "temp", "OCR")  

os.makedirs(ocr_folder, exist_ok=True)

df = pd.read_excel(excel_file, sheet_name="ALL", dtype=str) 

if "INV NO" not in df.columns:
    raise ValueError("Column 'INV NO' not found in the 'ALL' sheet of the Excel file.")

df["DOC NAME"] = df["INV NO"].astype(str) 

all_pdf_files = os.listdir(pdf_folder)

peb_numbers = []
inv_numbers = []
file_names = []
is_splitted_from = []
ok_flags = []

vam_pattern = re.compile(r"VAM-(\d+)(?:-(\d+))?")

for doc_name in df["DOC NAME"]:
    if doc_name in inv_numbers:  
        continue
    matching_files = [f for f in all_pdf_files if doc_name.lower() in f.lower()]

    doc_range = None
    range_start, range_end = None, None
    source_file = None

    for file_name in matching_files:
        range_match = vam_pattern.findall(file_name)  # Extract range from the actual file name
        if range_match and len(range_match[0]) == 2 and range_match[0][1]:  # Ensure it found both start and end
            if range_match:
                range_start = int(range_match[0][0])
                range_end = int(range_match[0][1]) if range_match[0][1] else range_start  # Handle single invoice case

            doc_range = file_name  # Assign the file name with range
            source_file = file_name
            break  # Stop after finding the first valid range


    if not matching_files:
        if doc_range:
            for num in range(range_start, range_end + 1):
                full_invoice_number = f"VAM-{num}"
                peb_numbers.append(peb_number if peb_number else None)  # Assign PEB number if exists
                inv_numbers.append(full_invoice_number)
                file_names.append(source_file)
                is_splitted_from.append(source_file)  # Set source file for range items
                ok_flags.append("OK" if source_file else "NO")

        else:
            peb_numbers.append(None)
            inv_numbers.append(doc_name)
            file_names.append(None)
            is_splitted_from.append(None)
            ok_flags.append("NO")

    else:
        pdf_path = os.path.join(pdf_folder, matching_files[0])
        # file_names.append(matching_files[0])

        print(f"\n🔍 Processing: {pdf_path}")

        ocr_pdf_path = os.path.join(ocr_folder, matching_files[0])

        with pdfplumber.open(pdf_path) as pdf:
            page_text = "\n".join([page.extract_text() or "" for page in pdf.pages])

        if page_text.strip():
            print("✅ PDF already contains text. Skipping OCR.")
            text_source = pdf_path  
        else:
            if not os.path.exists(ocr_pdf_path):
                print(f"⚠️ No OCR file found, running OCR on {pdf_path}...")
                os.system(f'python -m ocrmypdf --skip-text "{pdf_path}" "{ocr_pdf_path}"')

            text_source = ocr_pdf_path

        with pdfplumber.open(text_source) as pdf:
            relevant_pages = []
            for i, page in enumerate(pdf.pages):
                text = page.extract_text() or ""
                if re.search(r'\bBC\s*3\.?0\b.*PEMBERITAHUAN EKSPOR BARANG', text, re.IGNORECASE):
                    relevant_pages.append(f"\n📄 Page {i+1}:\n{text}")

        if relevant_pages:
            print("\n📝 OCR Extracted Text (Filtered for Relevant Pages):")
            # print("\n".join(relevant_pages))  
            # print("\n" + "-"*80 + "\n")
            page_text = "\n".join(relevant_pages)  
        else:
            print(f"⚠️ No relevant pages found in {text_source}.")
            peb_numbers.append(None)
            inv_numbers.append(doc_name)
            file_names.append(matching_files[0])
            is_splitted_from.append(None)
            ok_flags.append("NO")
            continue

        peb_number = None
        for line in page_text.split("\n"):
            if "Nomor Pendaftaran" in line:
                peb_number = line.split(":")[-1].strip()
                break

        doc_found = False
        for line in page_text.split("\n"): 
            if "Nomor & Tgl Invoice" in line and doc_name in line:
                doc_found = True
                break

            elif "Packing List" in line and doc_name in line:
                doc_found = True
                break

        if doc_range:
            doc_found = False
            for line in page_text.split("\n"):
                if "Packing List" in line:
                    for num in range(range_start, range_end + 1):
                        if f"VAM-{num}" in line:
                            doc_found = True
                            break

        if doc_range and doc_found:
            for num in range(range_start, range_end + 1):
                full_invoice_number = f"VAM-{num}"
                peb_numbers.append(peb_number)
                inv_numbers.append(full_invoice_number)
                file_names.append(source_file)
                is_splitted_from.append(source_file)
                ok_flags.append("OK")
        else:
            peb_numbers.append(peb_number)
            inv_numbers.append(doc_name)
            file_names.append(matching_files[0]) 
            is_splitted_from.append(None)
            ok_flags.append("OK" if doc_found else "NO")

        print(f"🔍 Lengths: PEB={len(peb_numbers)}, INV={len(inv_numbers)}, FILE={len(file_names)}, SPLIT={len(is_splitted_from)}, OK={len(ok_flags)}")
        if len(file_names) != len(peb_numbers):
            exit()

output_df = pd.DataFrame({
    # "PEB": peb_numbers,
    "INV": inv_numbers,
    "FILENAME (IF EXIST)": file_names,
    "isSplittedFrom": is_splitted_from,
    "OK?": ok_flags
})

output_file = os.path.join(root_dir, "temp", "Updated_CEK_LIST_PEB.xlsx")

with pd.ExcelWriter(output_file, mode="w", engine="xlsxwriter") as writer:
    output_df.to_excel(writer, sheet_name="Checked Data", index=False)