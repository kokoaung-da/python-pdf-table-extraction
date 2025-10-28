import os
import pdfplumber
import pandas as pd

# Input PDF file path
pdf_path = r"C:\pdf_query\Agent-List-updated-Jan.pdf"   # Replace with your PDF filename

# Output Excel file path (change this to any directory you want)
output_dir = r"C:\pdf_output"   # Replace with your desired folder
os.makedirs(output_dir, exist_ok=True)  # Create folder if not exists
output_excel = os.path.join(output_dir, "Agent_List_Extracted.xlsx")

# Collect all tables from all pages
all_tables = []

with pdfplumber.open(pdf_path) as pdf:
    for page_number, page in enumerate(pdf.pages, start=1):
        print(f"Processing page {page_number}...")
        tables = page.extract_tables()

        if not tables:
            print(f"⚠️ No tables found on page {page_number}")
            continue

        for table in tables:
            # Skip empty tables
            if not table or all(len(row) == 0 for row in table):
                continue
            
            # Convert raw table to DataFrame
            df = pd.DataFrame(table)
            all_tables.append(df)

# Combine all extracted tables into one DataFrame
if all_tables:
    combined_df = pd.concat(all_tables, ignore_index=True)
    combined_df.to_excel(output_excel, index=False)
    print(f"✅ Tables extracted and saved to: {output_excel}")
else:
    print("⚠️ No tables found in the entire PDF.")
