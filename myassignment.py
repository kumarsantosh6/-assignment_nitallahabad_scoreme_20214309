import pdfplumber
import pandas as pd
import re
from openpyxl import Workbook

def extract_tables_from_pdf(pdf_path, excel_path):
    # Initialize Excel writer
    writer = pd.ExcelWriter(excel_path, engine='openpyxl')
    
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            # Extract text with positions
            text = page.extract_text()
            
            # Try structured table extraction first
            tables = page.extract_tables({
                "vertical_strategy": "text", 
                "horizontal_strategy": "text"
            })
            
            if tables:
                for j, table in enumerate(tables):
                    if table and len(table) > 1:  # Ensure table is not empty
                        df = pd.DataFrame(table[1:], columns=table[0])
                        sheet_name = f"Page_{i+1}_Table_{j+1}"
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                # Fall back to semi-structured parsing
                transactions = []
                for line in text.split('\n'):
                    # Example pattern for extracting transactions
                    match = re.match(r'(\d{2}-[A-Za-z]{3}-\d{4})\s+(.*?)\s+([\d,]+\.\d{2})\s*(Dr|Cr)?', line)
                    if match:
                        date, desc, amount, drcr = match.groups()
                        transactions.append({
                            'Date': date,
                            'Description': desc.strip(),
                            'Amount': float(amount.replace(',', '')),
                            'Type': drcr if drcr else ''
                        })
                
                if transactions:
                    df = pd.DataFrame(transactions)
                    sheet_name = f"Page_{i+1}_Transactions"
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    writer.close()
    print(f"Extraction completed. Data saved to {excel_path}")
    return excel_path

# Example usage
if __name__ == "__main__":
    input_pdf = "input.pdf"  # Change this to your actual PDF file
    output_excel = "output.xlsx"
    extract_tables_from_pdf(input_pdf, output_excel)
