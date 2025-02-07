import os
import re
import pdfplumber
import pandas as pd
import multiprocessing
import time
import sys

def process_pdf(doc_name, all_pdf_files_lower, pdf_folder, ocr_folder, vam_pattern, peb_numbers, inv_numbers, file_names, is_splitted_from, ok_flags):
    lower_doc_name = doc_name.lower() 
    if lower_doc_name in inv_numbers:
        return

    matching_files = [all_pdf_files_lower[f] for f in all_pdf_files_lower if lower_doc_name in f]
    
    doc_range = None
    range_start, range_end = None, None
    source_file = None

    for file_name in matching_files:
        range_match = vam_pattern.findall(file_name)
        if range_match and len(range_match[0]) == 2 and range_match[0][1]:
            range_start = int(range_match[0][0])
            range_end = int(range_match[0][1]) if range_match[0][1] else range_start
            doc_range = file_name
            source_file = file_name
            break  

    if not matching_files:
        if doc_range:
            for num in range(range_start, range_end + 1):
                full_invoice_number = f"VAM-{num}"
                peb_numbers.append(None)
                inv_numbers.append(full_invoice_number)
                file_names.append(source_file)
                is_splitted_from.append(source_file)
                ok_flags.append("OK" if source_file else "NO")
        else:
            peb_numbers.append(None)
            inv_numbers.append(doc_name)
            file_names.append(None)
            is_splitted_from.append(None)
            ok_flags.append("NO")
        return

    pdf_path = os.path.join(pdf_folder, matching_files[0])
    print(f"\n🔍 Processing: {pdf_path}")

    ocr_pdf_path = os.path.join(ocr_folder, matching_files[0])
    text_source = pdf_path

    while True:  
        with pdfplumber.open(text_source) as pdf:
            page_text = "\n".join([page.extract_text() or "" for page in pdf.pages])

        if not page_text.strip():
            if not os.path.exists(ocr_pdf_path):
                os.system(f'python -m ocrmypdf --skip-text "{pdf_path}" "{ocr_pdf_path}"')
            text_source = ocr_pdf_path  
            continue  

        with pdfplumber.open(text_source) as pdf:
            relevant_pages = []
            for i, page in enumerate(pdf.pages):
                text = page.extract_text() or ""
                if re.search(r'\bBC\s*3\.?0\b.*PEMBERITAHUAN EKSPOR BARANG', text, re.IGNORECASE):
                    relevant_pages.append(f"\n📄 Page {i+1}:\n{text}")

        if relevant_pages:
            page_text = "\n".join(relevant_pages)
        else:
            peb_numbers.append(None)
            inv_numbers.append(doc_name)
            file_names.append(matching_files[0])
            is_splitted_from.append(None)
            ok_flags.append("NO")

            if source_file and ok_flags[-1] == "NO":
                os.system(f'python -m ocrmypdf --skip-text --deskew --rotate-pages "{pdf_path}" "{ocr_pdf_path}"')
                text_source = ocr_pdf_path  
                continue 

            return 
        break  

    peb_number = None
    for line in page_text.split("\n"):
        if "nomor pendaftaran" in line.lower():
            peb_number = line.split(":")[-1].strip()
            break

    doc_found = False
    lower_page_text = page_text.lower().split("\n")  

    for line in lower_page_text:
        if "nomor & tgl invoice" in line and lower_doc_name in line:
            doc_found = True
            break
        elif "packing list" in line and lower_doc_name in line:
            doc_found = True
            break

    if doc_range:
        doc_found = False
        for line in lower_page_text:
            if "nomor & tgl invoice" in line and lower_doc_name in line:
                if any(f"vam-{num}" in line for num in range(range_start, range_end + 1)):  
                    doc_found = True
                    break
            elif "packing list" in line:
                if any(f"vam-{num}" in line for num in range(range_start, range_end + 1)):
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

    print(f"🔍 Files Checked ={len(file_names)}")

if __name__ == "__main__":
    start_time = time.time()

    # Get arguments from main.py
    if len(sys.argv) != 4:
        print("❌ Error: Incorrect number of arguments. Usage: python pdfpl.py <excel_file> <pdf_folder> <output_file>")
        exit(1)

    excel_file, pdf_folder, output_file = sys.argv[1], sys.argv[2], sys.argv[3]
    
    root_dir = os.path.dirname(__file__)  
    ocr_folder = os.path.join(root_dir, "temp", "OCR")  
    os.makedirs(ocr_folder, exist_ok=True)

    df = pd.read_excel(excel_file, sheet_name="ALL", dtype=str) 

    if "INV NO" not in df.columns:
        raise ValueError("Column 'INV NO' not found in the 'ALL' sheet of the Excel file.")

    df["DOC NAME"] = df["INV NO"].astype(str) 
    all_pdf_files = os.listdir(pdf_folder)
    all_pdf_files_lower = {f.lower(): f for f in all_pdf_files}  
    vam_pattern = re.compile(r"VAM-(\d+)(?:-(\d+))?")

    manager = multiprocessing.Manager()
    peb_numbers = manager.list()
    inv_numbers = manager.list()
    file_names = manager.list()
    is_splitted_from = manager.list()
    ok_flags = manager.list()

    num_processes = multiprocessing.cpu_count()  
    pool = multiprocessing.Pool(processes=num_processes)
    pool.starmap(process_pdf, [(doc_name, all_pdf_files_lower, pdf_folder, ocr_folder, vam_pattern, peb_numbers, inv_numbers, file_names, is_splitted_from, ok_flags) for doc_name in df["DOC NAME"]])
    pool.close()
    pool.join()

    output_df = pd.DataFrame({
        "INV": list(inv_numbers),
        "FILENAME (IF EXIST)": list(file_names),
        "isSplittedFrom": list(is_splitted_from),
        "OK?": list(ok_flags)
    })

    output_df.to_excel(output_file, sheet_name="Checked Data", index=False)
    print(f"\n✅ Processing complete. Results saved to {output_file}.")
