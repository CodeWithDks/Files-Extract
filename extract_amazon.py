import pdfplumber
import pandas as pd
import os
import re
from datetime import datetime
import sys
import warnings
from typing import Dict, List, Optional, Tuple

# Suppress warnings
warnings.filterwarnings("ignore", category=UserWarning)

# CONFIGURATION 
def get_script_directory():
    """Get the directory where the script is located"""
    return os.path.dirname(os.path.abspath(sys.argv[0]))

BASE_DIR = os.path.join(get_script_directory(), "InvoiceDataExtraction")
INPUT_FOLDER = os.path.join(BASE_DIR, "Input_Folder")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "Output_Folder")

# Create directories if they don't exist
os.makedirs(INPUT_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

class InvoiceProcessor:
    @staticmethod
    def clean_text(text: str) -> str:
        """Normalize text for reliable parsing"""
        if not text:
            return ""
        text = re.sub(r'\s+', ' ', text)  # Replace multiple spaces with single space
        text = text.replace('\r\n', '\n').replace('\r', '\n')
        return text.strip()

    @staticmethod
    def clean_currency(amount_str: str) -> str:
        """Extract numeric value from currency string"""
        if not amount_str:
            return "0.00"
        # Remove currency symbols and extract numbers
        cleaned = re.sub(r'[^\d.,]', '', str(amount_str))
        # Handle Indian number format (₹1,234.56)
        cleaned = re.sub(r',', '', cleaned)
        return cleaned if cleaned else "0.00"

    @staticmethod
    def standardize_date(date_str: str) -> str:
        """Convert various date formats to standard DD/MM/YYYY"""
        if not date_str or date_str == "N/A":
            return "N/A"
        
        # Try different date patterns
        patterns = [
            r'(\d{2})\.(\d{2})\.(\d{4})',  # DD.MM.YYYY
            r'(\d{2})/(\d{2})/(\d{4})',   # DD/MM/YYYY  
            r'(\d{2})-(\d{2})-(\d{4})',   # DD-MM-YYYY
            r'(\d{4})-(\d{2})-(\d{2})',   # YYYY-MM-DD
        ]
        
        for pattern in patterns:
            match = re.search(pattern, date_str)
            if match:
                if len(match.group(1)) == 4:  # YYYY-MM-DD format
                    return f"{match.group(3)}/{match.group(2)}/{match.group(1)}"
                else:  # DD.MM.YYYY or DD/MM/YYYY format
                    return f"{match.group(1)}/{match.group(2)}/{match.group(3)}"
        
        return date_str

    @staticmethod
    def extract_amazon_invoice(text: str, filename: str) -> Dict[str, str]:
        """Enhanced Amazon invoice parser with robust pattern matching"""
        data = {
            "File Name": filename,
            "Invoice Source": "Amazon",
            "Order Number": "N/A",
            "Invoice Number": "N/A", 
            "Order Date": "N/A",
            "Invoice Date": "N/A",
            "Product Name": "N/A",
            "HSN Code": "N/A",
            "Quantity": "1",
            "Unit Price": "0.00",
            "Net Amount": "0.00",
            "Tax Rate": "N/A",
            "Tax Amount": "0.00",
            "Total Amount": "0.00",
            "Shipping Charges": "0.00",
            "Grand Total": "0.00",
            "Payment Mode": "N/A",
            "Seller Name": "N/A",
            "Seller GST": "N/A",
            "Billing Address": "N/A",
            "Shipping Address": "N/A"
        }

        try:
            # ORDER NUMBER
            order_patterns = [
                r"Order\s*Number[:\s]*([0-9\-]+)",
                r"Order\s*ID[:\s]*([0-9\-]+)",
                r"Amazon\s*Order[:\s]*([0-9\-]+)"
            ]
            for pattern in order_patterns:
                match = re.search(pattern, text, re.IGNORECASE)
                if match:
                    data["Order Number"] = match.group(1).strip()
                    break

            # INVOICE NUMBER  
            invoice_match = re.search(r"Invoice\s*Number[:\s]*([A-Z0-9\-]+)", text, re.IGNORECASE)
            if invoice_match:
                data["Invoice Number"] = invoice_match.group(1).strip()

            # DATES
            order_date_match = re.search(r"Order\s*Date[:\s]*([0-9./\-]+)", text, re.IGNORECASE)
            if order_date_match:
                data["Order Date"] = InvoiceProcessor.standardize_date(order_date_match.group(1).strip())

            invoice_date_match = re.search(r"Invoice\s*Date[:\s]*([0-9./\-]+)", text, re.IGNORECASE)
            if invoice_date_match:
                data["Invoice Date"] = InvoiceProcessor.standardize_date(invoice_date_match.group(1).strip())

            # PRODUCT DETAILS - Multi-line extraction
            lines = text.split('\n')
            for i, line in enumerate(lines):
                if re.search(r'^\s*1\s+', line):  # Line starting with "1 " (item number)
                    # Extract product name (could span multiple lines)
                    product_parts = []
                    current_line = line
                    
                    # Clean the line to get product name
                    product_line = re.sub(r'^\s*1\s+', '', current_line)  # Remove item number
                    product_parts.append(product_line)
                    
                    # Check next few lines for continuation
                    j = i + 1
                    while j < len(lines) and j < i + 5:  # Check up to 5 lines ahead
                        next_line = lines[j].strip()
                        if next_line and not re.match(r'^(HSN|₹|\d+%|IGST|Shipping)', next_line):
                            if not any(keyword in next_line.lower() for keyword in ['total', 'amount', 'tax', 'gst']):
                                product_parts.append(next_line)
                        else:
                            break
                        j += 1
                    
                    # Combine and clean product name
                    full_product = ' '.join(product_parts)
                    # Remove price information from product name
                    full_product = re.sub(r'₹[\d,.]+ \d+ ₹[\d,.]+.*$', '', full_product)
                    data["Product Name"] = full_product.strip()[:200]  # Limit length
                    break

            # HSN CODE
            hsn_match = re.search(r"HSN[:\s]*([0-9]+)", text, re.IGNORECASE)
            if hsn_match:
                data["HSN Code"] = hsn_match.group(1).strip()

            # PRICING DETAILS - Extract from table structure
            # Look for unit price pattern
            unit_price_match = re.search(r"₹([\d,]+\.?\d*)\s+1\s+₹", text)
            if unit_price_match:
                data["Unit Price"] = InvoiceProcessor.clean_currency(unit_price_match.group(1))

            # QUANTITY - usually 1 for most Amazon orders
            qty_match = re.search(r"₹[\d,]+\.?\d*\s+(\d+)\s+₹", text)
            if qty_match:
                data["Quantity"] = qty_match.group(1)

            # NET AMOUNT
            net_amount_match = re.search(r"₹[\d,]+\.?\d*\s+\d+\s+₹([\d,]+\.?\d*)", text)
            if net_amount_match:
                data["Net Amount"] = InvoiceProcessor.clean_currency(net_amount_match.group(1))

            # TAX DETAILS
            tax_rate_match = re.search(r"(\d+)%\s+IGST", text)
            if tax_rate_match:
                data["Tax Rate"] = f"{tax_rate_match.group(1)}%"

            tax_amount_match = re.search(r"IGST\s+₹([\d,]+\.?\d*)", text)
            if tax_amount_match:
                data["Tax Amount"] = InvoiceProcessor.clean_currency(tax_amount_match.group(1))

            # TOTAL AMOUNT (for main item)
            total_amount_match = re.search(r"IGST\s+₹[\d,]+\.?\d*\s+₹([\d,]+\.?\d*)", text)
            if total_amount_match:
                data["Total Amount"] = InvoiceProcessor.clean_currency(total_amount_match.group(1))

            # SHIPPING CHARGES
            shipping_match = re.search(r"Shipping\s+Charges\s+₹([\d,]+\.?\d*)", text, re.IGNORECASE)
            if shipping_match:
                data["Shipping Charges"] = InvoiceProcessor.clean_currency(shipping_match.group(1))

            # GRAND TOTAL
            grand_total_patterns = [
                r"TOTAL[:\s]*₹([\d,]+\.?\d*)",
                r"Invoice\s*Value[:\s]*₹?([\d,]+\.?\d*)",
                r"Grand\s*Total[:\s]*₹([\d,]+\.?\d*)"
            ]
            for pattern in grand_total_patterns:
                match = re.search(pattern, text, re.IGNORECASE)
                if match:
                    data["Grand Total"] = InvoiceProcessor.clean_currency(match.group(1))
                    break

            # PAYMENT MODE
            payment_match = re.search(r"Mode\s*of\s*Payment[:\s]*([^\n]+)", text, re.IGNORECASE)
            if payment_match:
                data["Payment Mode"] = payment_match.group(1).strip()

            # SELLER DETAILS
            seller_match = re.search(r"Sold\s*By[:\s]*([^\n*]+)", text, re.IGNORECASE)
            if seller_match:
                seller_name = seller_match.group(1).strip()
                # Clean seller name
                seller_name = re.sub(r'\*.*$', '', seller_name).strip()
                data["Seller Name"] = seller_name[:100]

            # SELLER GST
            gst_match = re.search(r"GST\s*Registration\s*No[:\s]*([A-Z0-9]+)", text, re.IGNORECASE)
            if gst_match:
                data["Seller GST"] = gst_match.group(1).strip()

            # BILLING ADDRESS
            billing_match = re.search(r"Billing\s*Address[:\s]*([^:]+?)(?=Shipping\s*Address|State/UT|$)", text, re.IGNORECASE | re.DOTALL)
            if billing_match:
                billing_addr = billing_match.group(1).strip()
                billing_addr = re.sub(r'\s+', ' ', billing_addr)
                data["Billing Address"] = billing_addr[:150]

            # SHIPPING ADDRESS  
            shipping_match = re.search(r"Shipping\s*Address[:\s]*([^:]+?)(?=Place\s*of|State/UT|$)", text, re.IGNORECASE | re.DOTALL)
            if shipping_match:
                shipping_addr = shipping_match.group(1).strip()
                shipping_addr = re.sub(r'\s+', ' ', shipping_addr)
                data["Shipping Address"] = shipping_addr[:150]

        except Exception as e:
            print(f" Error processing Amazon invoice {filename}: {str(e)}")

        return data

    @staticmethod
    def extract_flipkart_invoice(text: str, filename: str) -> Dict[str, str]:
        """Enhanced Flipkart invoice parser"""
        data = {
            "File Name": filename,
            "Invoice Source": "Flipkart", 
            "Order Number": "N/A",
            "Invoice Number": "N/A",
            "Order Date": "N/A",
            "Invoice Date": "N/A",
            "Product Name": "N/A",
            "HSN Code": "N/A",
            "Quantity": "1",
            "Unit Price": "0.00",
            "Net Amount": "0.00", 
            "Tax Rate": "N/A",
            "Tax Amount": "0.00",
            "Total Amount": "0.00",
            "Shipping Charges": "0.00",
            "Grand Total": "0.00",
            "Payment Mode": "N/A",
            "Seller Name": "N/A",
            "Seller GST": "N/A",
            "Billing Address": "N/A",
            "Shipping Address": "N/A"
        }

        try:
            # ORDER NUMBER
            order_match = re.search(r"Order\s*(?:ID|Number)[:\s]*([A-Z0-9]+)", text, re.IGNORECASE)
            if order_match:
                data["Order Number"] = order_match.group(1).strip()

            # INVOICE NUMBER
            invoice_patterns = [
                r"Invoice\s*(?:Number|No)[:\s#]*([A-Z0-9]+)",
                r"#\s*([A-Z0-9]{10,})"
            ]
            for pattern in invoice_patterns:
                match = re.search(pattern, text, re.IGNORECASE)
                if match:
                    data["Invoice Number"] = match.group(1).strip()
                    break

            # DATES
            order_date_match = re.search(r"Order\s*Date[:\s]*([0-9.\-/]+)", text, re.IGNORECASE)
            if order_date_match:
                data["Order Date"] = InvoiceProcessor.standardize_date(order_date_match.group(1).strip())

            invoice_date_match = re.search(r"Invoice\s*Date[:\s]*([0-9.\-/]+)", text, re.IGNORECASE)
            if invoice_date_match:
                data["Invoice Date"] = InvoiceProcessor.standardize_date(invoice_date_match.group(1).strip())

            # PRODUCT NAME - Multiple approaches
            product_patterns = [
                r"Product\s*Description\s+Qty.*?\n([^\n]+)",
                r"Description\s+Qty.*?\n([^\n]+)",
                r"([A-Za-z].*?)\s+HSN[:\s]*\d+",
                r"Ordered\s*Through.*?\n([^\n]+)"
            ]
            
            for pattern in product_patterns:
                match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
                if match:
                    product_name = match.group(1).strip()
                    # Clean product name
                    product_name = re.sub(r'\d+\s+[\d,]+\.?\d*.*$', '', product_name)  # Remove trailing numbers/prices
                    product_name = re.sub(r'^\d+\s+', '', product_name)  # Remove leading numbers
                    if len(product_name) > 10:  # Valid product name
                        data["Product Name"] = product_name[:200]
                        break

            # HSN CODE
            hsn_match = re.search(r"HSN[:\s]*(\d+)", text, re.IGNORECASE)
            if hsn_match:
                data["HSN Code"] = hsn_match.group(1).strip()

            # QUANTITY
            qty_match = re.search(r"Qty\s+(\d+)", text, re.IGNORECASE)
            if qty_match:
                data["Quantity"] = qty_match.group(1)

            # PRICING - Look for table structure
            # Gross Amount (Unit Price)
            gross_amount_match = re.search(r"Gross\s*Amount[^₹]*₹?\s*([\d,]+\.?\d*)", text, re.IGNORECASE)
            if gross_amount_match:
                data["Unit Price"] = InvoiceProcessor.clean_currency(gross_amount_match.group(1))

            # Taxable Value (Net Amount)
            taxable_match = re.search(r"Taxable\s*[Vv]alue[^₹]*₹?\s*([\d,]+\.?\d*)", text, re.IGNORECASE)
            if taxable_match:
                data["Net Amount"] = InvoiceProcessor.clean_currency(taxable_match.group(1))

            # TAX DETAILS
            tax_rate_match = re.search(r"(\d+\.?\d*)%\s*IGST", text)
            if tax_rate_match:
                data["Tax Rate"] = f"{tax_rate_match.group(1)}%"

            tax_amount_match = re.search(r"IGST[^₹]*₹?\s*([\d,]+\.?\d*)", text, re.IGNORECASE)
            if tax_amount_match:
                data["Tax Amount"] = InvoiceProcessor.clean_currency(tax_amount_match.group(1))

            # TOTAL AMOUNT
            total_patterns = [
                r"Total[^₹]*₹?\s*([\d,]+\.?\d*)",
                r"Grand\s*Total[^₹]*₹?\s*([\d,]+\.?\d*)"
            ]
            for pattern in total_patterns:
                match = re.search(pattern, text, re.IGNORECASE)
                if match:
                    data["Grand Total"] = InvoiceProcessor.clean_currency(match.group(1))
                    break

            # SHIPPING CHARGES
            shipping_patterns = [
                r"Shipping\s*(?:and\s*)?(?:Handling\s*)?Charges[^₹]*₹?\s*([\d,]+\.?\d*)",
                r"Delivery\s*Charges[^₹]*₹?\s*([\d,]+\.?\d*)"
            ]
            for pattern in shipping_patterns:
                match = re.search(pattern, text, re.IGNORECASE)
                if match and match.group(1) != "0.00":
                    data["Shipping Charges"] = InvoiceProcessor.clean_currency(match.group(1))
                    break

            # SELLER DETAILS
            seller_match = re.search(r"Sold\s*By[:\s]*([^,\n]+)", text, re.IGNORECASE)
            if seller_match:
                data["Seller Name"] = seller_match.group(1).strip()[:100]

            # SELLER GST
            gst_patterns = [
                r"GSTIN[:\s-]*([A-Z0-9]{15})",
                r"GST[:\s]*([A-Z0-9]{15})"
            ]
            for pattern in gst_patterns:
                match = re.search(pattern, text, re.IGNORECASE)
                if match:
                    data["Seller GST"] = match.group(1).strip()
                    break

            # ADDRESSES
            billing_match = re.search(r"Bill\s*To\s*([^:]+?)(?=Ship\s*To|Order\s*ID|$)", text, re.IGNORECASE | re.DOTALL)
            if billing_match:
                billing_addr = billing_match.group(1).strip()
                billing_addr = re.sub(r'\s+', ' ', billing_addr)
                data["Billing Address"] = billing_addr[:150]

            shipping_match = re.search(r"Ship\s*To\s*([^:]+?)(?=Bill\s*To|Order\s*ID|$)", text, re.IGNORECASE | re.DOTALL)
            if shipping_match:
                shipping_addr = shipping_match.group(1).strip()
                shipping_addr = re.sub(r'\s+', ' ', shipping_addr)
                data["Shipping Address"] = shipping_addr[:150]

        except Exception as e:
            print(f" Error processing Flipkart invoice {filename}: {str(e)}")

        return data

    @staticmethod
    def detect_invoice_type(text: str, filename: str) -> str:
        """Detect invoice type based on content and filename"""
        text_lower = text.lower()
        filename_lower = filename.lower()
        
        # Check filename first
        if 'amazon' in filename_lower:
            return 'amazon'
        elif 'flipkart' in filename_lower:
            return 'flipkart'
        
        # Check content
        if 'amazon' in text_lower or 'ASSPL-Amazon' in text:
            return 'amazon'
        elif 'flipkart' in text_lower:
            return 'flipkart'
        
        return 'unknown'

def process_invoices():
    """Main processing function"""
    print(f"\n{'='*60}")
    print(" ENHANCED INVOICE DATA EXTRACTION SYSTEM")
    print(f"{'='*60}")
    print(f" Input Folder: {INPUT_FOLDER}")
    print(f" Output Folder: {OUTPUT_FOLDER}\n")

    all_invoices = []
    processed_files = 0
    failed_files = []

    try:
        # Get all PDF files
        files = [f for f in os.listdir(INPUT_FOLDER) if f.lower().endswith(".pdf")]
        if not files:
            print(" No PDF files found in input folder")
            print(f"Please place PDF files in: {INPUT_FOLDER}")
            return

        print(f" Found {len(files)} PDF file(s) to process:")
        print("-" * 40)
        
        for filename in files:
            filepath = os.path.join(INPUT_FOLDER, filename)
            print(f" Processing: {filename}")
            
            try:
                with pdfplumber.open(filepath) as pdf:
                    # Extract text from all pages
                    text_parts = []
                    for page in pdf.pages:
                        page_text = page.extract_text()
                        if page_text:
                            text_parts.append(page_text)
                    
                    full_text = "\n".join(text_parts)
                    clean_text = InvoiceProcessor.clean_text(full_text)
                    
                    if not clean_text:
                        print(f"    No text extracted from {filename}")
                        failed_files.append(filename)
                        continue
                    
                    # Detect invoice type and extract data
                    invoice_type = InvoiceProcessor.detect_invoice_type(clean_text, filename)
                    
                    if invoice_type == 'amazon':
                        result = InvoiceProcessor.extract_amazon_invoice(clean_text, filename)
                        print(f"    Amazon invoice processed")
                    elif invoice_type == 'flipkart':
                        result = InvoiceProcessor.extract_flipkart_invoice(clean_text, filename)
                        print(f"    Flipkart invoice processed")
                    else:
                        print(f"    Unknown invoice type for {filename}")
                        # Still try to extract as generic
                        result = InvoiceProcessor.extract_amazon_invoice(clean_text, filename)
                        result["Invoice Source"] = "Unknown"
                    
                    all_invoices.append(result)
                    processed_files += 1

            except Exception as e:
                failed_files.append(filename)
                print(f"    Failed: {str(e)}")

        # Save results
        if all_invoices:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = os.path.join(OUTPUT_FOLDER, f"extracted_invoices_{timestamp}.xlsx")
            
            # Create DataFrame
            df = pd.DataFrame(all_invoices)
            
            # Reorder columns for better readability
            column_order = [
                "File Name", "Invoice Source", "Order Number", "Invoice Number",
                "Order Date", "Invoice Date", "Product Name", "HSN Code", 
                "Quantity", "Unit Price", "Net Amount", "Tax Rate", "Tax Amount",
                "Total Amount", "Shipping Charges", "Grand Total", "Payment Mode",
                "Seller Name", "Seller GST", "Billing Address", "Shipping Address"
            ]
            
            df = df.reindex(columns=column_order)
            
            # Save to Excel with formatting
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Invoice_Data', index=False)
                
                # Auto-adjust column widths
                worksheet = writer.sheets['Invoice_Data']
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width

            print(f"\n{'='*60}")
            print(" PROCESSING SUMMARY")
            print(f"{'='*60}")
            print(f" Successfully processed: {processed_files} files")
            print(f" Failed to process: {len(failed_files)} files")
            
            if failed_files:
                print("\nFailed files:")
                for f in failed_files:
                    print(f"   - {f}")
            
            print(f"\n Output saved to:")
            print(f"   {os.path.basename(output_file)}")
            print(f"\n Total records extracted: {len(all_invoices)}")
            
            # Show sample data
            if len(all_invoices) > 0:
                print(f"\n Sample extracted data:")
                print("-" * 40)
                sample = all_invoices[0]
                for key, value in list(sample.items())[:8]:  # Show first 8 fields
                    print(f"   {key}: {value}")
                if len(sample) > 8:
                    print("   ... (and more fields)")
            
            print("\n Invoice extraction completed successfully!")
            
        else:
            print("\n No invoices were successfully processed")

    except Exception as e:
        print(f"\n Critical error: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    process_invoices()