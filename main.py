import os
import subprocess

def get_user_input(prompt, default_value):
    user_input = input(f"{prompt} (default: {default_value}): ").strip()
    return user_input if user_input else default_value

if __name__ == "__main__":
    root_dir = os.path.dirname(__file__)  # Get script directory

    # Get user inputs
    excel_filename = get_user_input("Enter the Excel file name (without .xlsx)", "CEK LIST PEB")
    pdf_folder_name = get_user_input("Enter the PDF folder name", "ALL2022")
    output_filename = get_user_input("Enter the output Excel file name (without .xlsx)", "Updated_CEK_LIST_PEB")

    # Construct file paths
    excel_file = os.path.join(root_dir, "temp", f"{excel_filename}.xlsx")
    pdf_folder = os.path.join(root_dir, "temp", pdf_folder_name)
    output_file = os.path.join(root_dir, "temp", f"{output_filename}.xlsx")

    # Ensure files and directories exist
    if not os.path.exists(excel_file):
        print(f"❌ Error: Excel file '{excel_file}' not found.")
        exit(1)

    if not os.path.exists(pdf_folder):
        print(f"❌ Error: PDF folder '{pdf_folder}' not found.")
        exit(1)

    # Run pdfpl.py as a subprocess
    subprocess.run(["python", "pdfpl.py", excel_file, pdf_folder, output_file])
