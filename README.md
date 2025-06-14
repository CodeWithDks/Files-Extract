# Invoice Data Extraction System

This project extracts structured data from Amazon and Flipkart invoice PDFs and generates a professional Excel report.  
It is created as part of the F-AI internship application process.

---

## 📌 Features

- 📄 Processes Amazon & Flipkart invoice PDFs
- 📊 Extracts 20+ fields, including:
  - Invoice Number, Order ID, Dates
  - Buyer & Seller Details
  - GST, Tax Type, Tax Amount
  - Product Description, Quantity, Unit Price
  - Total Amount, Payment Method, Shipping Address
- 📁 Generates timestamped Excel reports
- 🛡 Handles missing values with `"N/A"` placeholders

---

## 🛠 Setup Instructions

### 1. Clone the Repository

git clone https://github.com/CodeWithDks/Files-Extract.git
cd Files-Extract/InvoiceDataExtraction
2. Install Dependencies
bash

pip install -r requirements.txt
📁 Folder Structure

InvoiceDataExtraction/
├── Input_Folder/         # Invoice PDFs (Amazon, Flipkart)
├── Output_Folder/        # Extracted Excel file
├──Python script to run
├── requirements.txt      # Python libraries
└── README.md             # Project documentation
▶️ How to Run
Place PDFs in the Input_Folder/

Navigate to Coding_Folder/ and run:


python extract_invoice_data.py
Extracted Excel file will be saved in Output_Folder/

📊 Sample Output
Output file: extracted_invoices.xlsx
Contains structured data for each invoice in rows and columns.
