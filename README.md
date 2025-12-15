# Advanced PDF Order & Line-Item Extractor (Logistics Automation)

## 📌 Overview
This Python automation script extracts **structured order and line-item data** from complex logistics PDFs and converts it into a clean Excel report.

It is designed to handle **multi-page PDFs**, extract metadata, and accurately parse **SKU-wise quantities and descriptions** using positional text analysis.

---

## 🚀 Features
- Processes bulk PDF files from a folder
- Extracts fields such as:
  - System Number
  - Order Number
  - Tracking Number
- Parses SKU, description, and ordered quantity from tabular PDF layouts
- Uses positional (X/Y) coordinates for reliable extraction
- Automatically removes duplicate rows
- Supports optional file-processing limits
- Exports consolidated data to Excel

---

## 🛠️ Technologies Used
- Python
- pdfplumber
- Pandas
- Regular Expressions
- Positional text extraction (X/Y coordinates)

---

## 📂 Extracted Fields
- System No  
- Order No  
- Tracking Number 
- SKU  
- Description  
- QTY ORD  

---

## ⚙️ How Extraction Works
- Header fields are extracted once per PDF using regex
- Line items are detected by grouping words based on vertical alignment
- SKU, description, and quantity are identified using column-based X-axis ranges
- Quantity values are normalized and converted to integers

---

## ▶️ How to Run

#### 1️⃣ Install Dependencies
```bash
pip install -r requirements.txt
```

#### 2️⃣ Run the Script
```bash
python extract_2.py
```
---

## ▶️ Running via Batch File (Windows)

This project includes a Windows batch file for easy execution without manually typing commands.

### What the batch file does
- Prompts the user to optionally limit the number of PDFs to process
- Runs the Python extraction script using system Python
- Keeps the console open after execution for review

### Usage
1. Ensure Python and required packages are installed
2. Update folder paths inside the batch file
3. Double-click `run_extraction.bat`
4. Enter the number of PDFs to process or press **Enter** to process all files

---

## ⚠️ Requirements
- Python 3.8+
- PDFs must contain selectable text (not scanned images)

---

## 📈 Use Case

#### Ideal for:
- Logistics order processing
- Shipment and order reporting
- Replacing manual PDF data entry
- Warehouse and fulfillment analysis
- Automating PDF-to-Excel data extraction
