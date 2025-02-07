# PDF Processing Automation

## 📌 Overview
This project automates the extraction of invoice and PEB data from PDFs using `pdfplumber` and OCR. It reads an Excel file containing invoice numbers, scans PDFs in a specified folder, and outputs a processed Excel file.

---

## 🛠️ Setup Instructions

### 1️⃣ Install Dependencies
Before running the scripts, install the necessary dependencies.

#### **Windows**
```sh
pip install -r requirements.txt
```

#### **Mac**
```sh
pip3 install -r requirements.txt
```

#### **Linux**
```sh
sudo apt install python3-pip -y  # Ensure pip is installed
pip3 install -r requirements.txt
```

If `pip` is not installed, install it first:
- **Windows**: `python -m ensurepip --default-pip`
- **Mac**: `brew install python3`
- **Linux**: `sudo apt install python3-pip`

---

### 2️⃣ Install Poppler
**Poppler** is required for handling PDFs. Install it based on your operating system:

#### **Windows**
1. [Download Poppler for Windows](https://github.com/oschwartz10612/poppler-windows/releases/)
2. Extract the folder and add `poppler/bin` to your system PATH.
3. Restart your terminal.

#### **Mac**
```sh
brew install poppler
```

#### **Linux**
```sh
sudo apt install poppler-utils
```

---

### 3️⃣ Folder Structure
Ensure the following structure exists inside your project:
```
PDF-Reading-Document-Fitting/
│── temp/
│   ├── CEK LIST PEB.xlsx  (Your input Excel file)
│   ├── ALL2022/           (Folder containing PDFs)
│   ├── OCR/               (Auto-created for OCR-processed PDFs)
│── poppler/               (Poppler folder, now included in the repo)
│── main.py                (Entry point to run the script)
│── pdfpl.py               (Processing script)
│── requirements.txt        (Dependencies list)
│── README.md              (This file)
```

---

### 4️⃣ Running the Script
Navigate to the project folder in the terminal and run:

#### **Windows**
```sh
python main.py
```

#### **Mac**
```sh
python3 main.py
```

#### **Linux**
```sh
python3 main.py
```

The script will prompt you to enter:
- **Excel file name** (without `.xlsx`)
- **PDF folder name**
- **Output file name** (without `.xlsx`)

Example Input:
```
Enter the Excel file name (without .xlsx): File PEB
Enter the PDF folder name: ALL2022
Enter the output Excel file name (without .xlsx): Processed_Result
```

Example Output File: `temp/Processed_Result.xlsx`

---

### 5️⃣ Error Handling
- If the script fails to find the Excel file or PDF folder, it will exit with an error.
- Ensure the `temp/` folder contains the required files before running the script.

---

## 👨‍💻 Developed by
**Favian Izza Diasputra**