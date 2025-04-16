#!/usr/bin/env python3
"""
Excel Data Extractor - Command Line Version
This script extracts data from Excel files in a ZIP archive and
merges selected columns into a new Excel file.

This is a command-line version of the application. The full PyQt5 GUI
version (excel_extractor_qt.py) can be run on MacOS systems with PyQt5 installed.
"""

import os
import sys
import tempfile
import pandas as pd
from zipfile import ZipFile
from pathlib import Path
import xlwt
import argparse

def extract_zip_file(zip_path, extract_dir):
    """
    Extract Excel files from a ZIP archive
    
    Parameters:
    - zip_path: Path to the ZIP file
    - extract_dir: Directory to extract files to
    
    Returns:
    - A list of paths to extracted Excel files
    """
    excel_files = []
    
    try:
        print(f"Opening ZIP file: {zip_path}")
        
        with ZipFile(zip_path, 'r') as zip_ref:
            # List all files in the ZIP
            file_list = zip_ref.namelist()
            
            print(f"Found {len(file_list)} files in ZIP archive")
            
            # Extract only Excel files
            for file_name in file_list:
                lower_name = file_name.lower()
                if lower_name.endswith('.xlsx') or lower_name.endswith('.xls'):
                    # Handle folder paths in ZIP
                    if file_name.endswith('/') or os.path.basename(file_name) == '':
                        continue
                        
                    # Extract the file
                    try:
                        print(f"Extracting: {file_name}")
                        zip_ref.extract(file_name, extract_dir)
                        full_path = os.path.join(extract_dir, file_name)
                        excel_files.append(full_path)
                    except Exception as extract_error:
                        print(f"Could not extract {file_name}: {str(extract_error)}")
                
            # Also look for Excel files in any folders that were extracted
            for root, dirs, files in os.walk(extract_dir):
                for file in files:
                    if file.lower().endswith(('.xlsx', '.xls')) and os.path.join(root, file) not in excel_files:
                        excel_files.append(os.path.join(root, file))
                        print(f"Found additional Excel file: {file}")
    
    except Exception as e:
        print(f"Error extracting ZIP file: {str(e)}")
        return []
    
    print(f"Extracted {len(excel_files)} Excel files")
    return excel_files

def read_excel_files(file_paths):
    """
    Read data from multiple Excel files
    
    Parameters:
    - file_paths: List of paths to Excel files
    
    Returns:
    - A nested dictionary structure: {file_name: {sheet_name: dataframe}}
    """
    file_data = {}
    
    if not file_paths:
        print("No Excel files to process")
        return file_data
    
    print(f"Reading {len(file_paths)} Excel files...")
    
    for file_path in file_paths:
        try:
            # Get just the filename without path
            file_name = os.path.basename(file_path)
            print(f"Reading: {file_name}")
            
            # Read all sheets from the Excel file
            try:
                excel_file = pd.ExcelFile(file_path)
                sheet_names = excel_file.sheet_names
                print(f"Found {len(sheet_names)} sheets in {file_name}")
            except Exception as excel_error:
                print(f"Error opening Excel file '{file_name}': {str(excel_error)}")
                
                # Try alternate approach for older Excel formats
                try:
                    # For xls files
                    if file_path.lower().endswith('.xls'):
                        df = pd.read_excel(file_path, engine='xlrd')
                        file_data[file_name] = {"Sheet1": df}
                        print(f"Successfully read {file_name} using xlrd engine")
                        continue
                except Exception as alt_error:
                    print(f"Alternative read approach failed: {str(alt_error)}")
                continue
            
            # Initialize the entry for this file
            file_data[file_name] = {}
            
            # Read each sheet and store its data
            for sheet_name in sheet_names:
                try:
                    df = pd.read_excel(excel_file, sheet_name=sheet_name)
                    
                    # Only keep sheets that have data
                    if not df.empty:
                        file_data[file_name][sheet_name] = df
                        print(f"Sheet '{sheet_name}' has {len(df)} rows and {len(df.columns)} columns")
                    else:
                        print(f"Sheet '{sheet_name}' is empty, skipping")
                except Exception as e:
                    print(f"Error reading sheet '{sheet_name}': {str(e)}")
                    continue
            
            # If no sheets were successfully read, remove this file entry
            if not file_data[file_name]:
                print(f"No data found in file '{file_name}'")
                del file_data[file_name]
                
        except Exception as e:
            print(f"Error reading file '{os.path.basename(file_path)}': {str(e)}")
            continue
    
    # Provide summary
    file_count = len(file_data)
    if file_count > 0:
        sheet_count = sum(len(sheets) for sheets in file_data.values())
        print(f"Successfully read {file_count} files with a total of {sheet_count} sheets")
    else:
        print("Could not read any data from the Excel files")
    
    return file_data

def interactive_column_selection(file_data):
    """
    Allow the user to interactively select columns from each file and sheet
    
    Parameters:
    - file_data: Nested dictionary of file data {file_name: {sheet_name: dataframe}}
    
    Returns:
    - selected_columns: Nested dictionary of selected columns {file_name: {sheet_name: [columns]}}
    """
    selected_columns = {}
    
    for file_name, sheets in file_data.items():
        selected_columns[file_name] = {}
        
        print(f"\n===== File: {file_name} =====")
        
        for sheet_name, df in sheets.items():
            print(f"\n--- Sheet: {sheet_name} ---")
            print("Available columns:")
            
            # Print columns with numbers for selection
            columns = df.columns.tolist()
            for i, col in enumerate(columns):
                print(f"{i+1}. {col}")
            
            # Get user selection
            print("\nEnter column numbers to select (comma-separated), 'all' for all columns, or 'none' for none:")
            selection = input("> ").strip().lower()
            
            if selection == 'all':
                selected_columns[file_name][sheet_name] = columns
                print(f"Selected all {len(columns)} columns")
            elif selection == 'none':
                selected_columns[file_name][sheet_name] = []
                print("No columns selected")
            else:
                try:
                    # Parse comma-separated numbers
                    indices = [int(idx.strip()) - 1 for idx in selection.split(',') if idx.strip()]
                    # Filter valid indices
                    valid_indices = [idx for idx in indices if 0 <= idx < len(columns)]
                    
                    # Get the selected columns
                    selected_cols = [columns[idx] for idx in valid_indices]
                    selected_columns[file_name][sheet_name] = selected_cols
                    
                    print(f"Selected {len(selected_cols)} columns: {', '.join(selected_cols)}")
                except Exception as e:
                    print(f"Error parsing selection: {str(e)}")
                    selected_columns[file_name][sheet_name] = []
    
    # Calculate total selected columns
    total_selected = sum(len(cols) for file in selected_columns.values() for cols in file.values())
    print(f"\nTotal columns selected: {total_selected}")
    
    return selected_columns

def process_and_merge_data(file_data, selected_columns, output_path):
    """
    Process and merge selected data from multiple Excel files
    
    Parameters:
    - file_data: Nested dictionary of file data {file_name: {sheet_name: dataframe}}
    - selected_columns: Nested dictionary of selected columns {file_name: {sheet_name: [columns]}}
    - output_path: Path to save the merged Excel file
    
    Returns:
    - True if successful, False otherwise
    """
    try:
        print("Starting data processing...")
        
        # Create a new workbook
        workbook = xlwt.Workbook()
        
        # Track the number of worksheets created
        worksheet_count = 0
        
        # Process each file
        for file_name, sheets in file_data.items():
            print(f"Processing file: {file_name}")
            
            # Process each sheet in the file
            for sheet_name, df in sheets.items():
                # Get the selected columns for this sheet
                cols = selected_columns.get(file_name, {}).get(sheet_name, [])
                
                # Skip if no columns were selected for this sheet
                if not cols:
                    print(f"No columns selected for {file_name} - {sheet_name}, skipping")
                    continue
                
                print(f"Processing sheet: {sheet_name} with {len(cols)} selected columns")
                
                # Extract only the selected columns
                subset_df = df[cols].copy()
                
                # Create a worksheet name from the file and sheet names
                # Ensure it's valid and not too long for Excel
                ws_name = f"{Path(file_name).stem}_{sheet_name}"
                ws_name = ws_name.replace("[", "").replace("]", "").replace(":", "")
                ws_name = ws_name[:31]  # Excel has 31 char limit for sheet names
                
                # Handle duplicate sheet names by appending a number
                original_ws_name = ws_name
                counter = 1
                while ws_name in [sheet.name for sheet in workbook.get_sheets()]:
                    ws_name = f"{original_ws_name[:27]}_{counter}"
                    counter += 1
                
                # Create a new worksheet
                worksheet = workbook.add_sheet(ws_name)
                worksheet_count += 1
                
                # Write column headers
                for col_idx, col_name in enumerate(subset_df.columns):
                    worksheet.write(0, col_idx, col_name)
                
                # Write data rows
                for row_idx, row in enumerate(subset_df.values):
                    for col_idx, value in enumerate(row):
                        # Handle NaN values
                        if pd.isna(value):
                            worksheet.write(row_idx + 1, col_idx, "")
                        else:
                            worksheet.write(row_idx + 1, col_idx, value)
        
        # Create a summary sheet
        summary = workbook.add_sheet("Summary")
        
        # Write summary headers
        summary.write(0, 0, "File")
        summary.write(0, 1, "Sheet")
        summary.write(0, 2, "Columns Extracted")
        
        # Write summary data
        row = 1
        for file_name, sheets in selected_columns.items():
            for sheet_name, cols in sheets.items():
                if cols:  # Only include sheets where columns were selected
                    summary.write(row, 0, file_name)
                    summary.write(row, 1, sheet_name)
                    summary.write(row, 2, ", ".join(cols))
                    row += 1
        
        # Save the workbook
        print(f"Saving output to: {output_path}")
        workbook.save(output_path)
        
        print(f"Processing complete. Created {worksheet_count} worksheets plus summary.")
        return True
    
    except Exception as e:
        print(f"Error processing and merging data: {str(e)}")
        return False

def main():
    parser = argparse.ArgumentParser(description="Excel Data Extractor - Extract and merge data from Excel files in a ZIP archive")
    parser.add_argument("zip_file", help="Path to the ZIP file containing Excel files")
    parser.add_argument("output_file", help="Path to save the merged Excel file")
    
    args = parser.parse_args()
    
    # Check if the ZIP file exists
    if not os.path.exists(args.zip_file):
        print(f"Error: ZIP file does not exist: {args.zip_file}")
        return 1
    
    # Create a temporary directory for extraction
    temp_dir = tempfile.mkdtemp()
    print(f"Created temporary directory: {temp_dir}")
    
    try:
        # Extract Excel files from the ZIP
        extracted_files = extract_zip_file(args.zip_file, temp_dir)
        
        if not extracted_files:
            print("No Excel files found in the ZIP archive.")
            return 1
        
        # Read Excel files and store their data
        file_data = read_excel_files(extracted_files)
        
        if not file_data:
            print("Could not read any data from the Excel files.")
            return 1
        
        # Let the user select columns
        selected_columns = interactive_column_selection(file_data)
        
        # Calculate total selected columns
        total_selected = sum(len(cols) for file in selected_columns.values() for cols in file.values())
        
        if total_selected == 0:
            print("No columns selected. Nothing to process.")
            return 1
        
        # Process and merge the selected data
        success = process_and_merge_data(file_data, selected_columns, args.output_file)
        
        if success:
            print(f"\nMerged data successfully saved to: {args.output_file}")
            return 0
        else:
            print("\nFailed to process and merge data.")
            return 1
        
    finally:
        # Clean up the temporary directory
        if os.path.exists(temp_dir):
            try:
                import shutil
                shutil.rmtree(temp_dir)
                print(f"Cleaned up temporary directory: {temp_dir}")
            except Exception as e:
                print(f"Error cleaning up temporary directory: {e}")
    
if __name__ == "__main__":
    sys.exit(main())