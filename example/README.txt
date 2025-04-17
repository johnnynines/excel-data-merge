EXAMPLE DATA STRUCTURE

This folder would typically contain example Excel files packaged in a ZIP archive for testing the application.

To create a test file:

1. Create several Excel files with sample data
2. Select all of these files
3. Right-click and choose "Compress Items" on macOS to create a ZIP archive
4. Use this ZIP as your test input file for the Excel Data Extractor

Example Excel File Structure:
----------------------------
File: sales_data.xlsx
Sheets:
- 2023_Q1 (columns: Date, Product, Region, Amount)
- 2023_Q2 (columns: Date, Product, Region, Amount)

File: customer_info.xlsx
Sheets:
- Customers (columns: ID, Name, Email, Phone, Region)
- Contacts (columns: ID, Name, Position, Company)

When using the application, you can extract specific columns from each sheet,
such as extracting only the "Date", "Product", and "Amount" columns from the 
sales data while extracting only the "ID", "Name", and "Region" columns from 
the customer information.