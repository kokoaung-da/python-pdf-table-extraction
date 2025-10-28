# PDF Table Extractor (Python)

**Overview**

This Python script extracts tables from multi-page PDF files and compiles them into a single Excel workbook.

I built it to solve a common problem — Excel’s Power Query often struggles with very large PDFs, especially those containing hundreds or thousands of pages of tabular data. Using pdfplumber, the script reads each page, pulls out any tables it finds, and merges them into one clean Excel file.

**Features**

Extracts tables from every page of a PDF (tested successfully on files with over 1,000 pages)

- Combines all tables into one well-structured Excel sheet

- Skips empty or unreadable tables automatically

- Creates the output folder if it doesn’t already exist

- Displays progress messages as each page is processed

**Required Libraries**

- pdfplumber 
- pandas 
- openpyxl

**Pros**

Handles large PDFs that Power Query can’t process

Fully automated — just set the paths and run

Detects multiple tables per page

Lightweight and easy to modify

Outputs directly to .xlsx format

