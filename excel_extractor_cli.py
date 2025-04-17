#!/usr/bin/env python3
"""
Excel Data Extractor - Command Line Version
This script extracts data from Excel files in a ZIP archive and
merges selected columns into a new Excel file.

This is a command-line version of the application. The full wxPython GUI
version (excel_extractor_wx.py) can be run on MacOS systems with wxPython installed.
"""

import os
import sys
import tempfile
import shutil
import pandas as pd
from zipfile import ZipFile
import xlwt
from pathlib import Path
import argparse

# Import shared functionality
from file_processor import extract_zip_file, read_excel_files, process_and_merge_data

def interactive_column_selection(file_data):
    """
    Allow the user to interactively select columns from each file and sheet
    
    Parameters:
    - file_data: Nested dictionary of file data {file_name: {sheet_name: dataframe}}
    
    Returns:
    - selected_columns: Nested dictionary of selected columns {file_name: {sheet_name: [columns]}}
    """
    selected_columns = {}
    
    print("\n=== COLUMN SELECTION ===")
    print("Select columns to extract from each file and sheet")
    
    for file_name, sheets in file_data.items():
        selected_columns[file_name] = {}
        
        print(f"\nFILE: {file_name}")
        
        for sheet_name, df in sheets.items():
            selected_columns[file_name][sheet_name] = []
            
            print(f"\n  SHEET: {sheet_name}")
            print(f"  Total columns: {len(df.columns)}")
            
            # Display data preview
            print("\n  DATA PREVIEW (first 3 rows):")
            preview = df.head(3).to_string()
            for line in preview.split('\n'):
                print(f"  {line}")
            
            # Get descriptive column names for better display
            try:
                from file_processor import detect_descriptive_column_names
                descriptive_names = detect_descriptive_column_names(df)
                
                # Display column options with descriptive names
                print("\n  Available columns:")
                for i, col in enumerate(df.columns):
                    desc_name = descriptive_names.get(col, col)
                    if desc_name != col and not col.startswith("Column_"):
                        print(f"  {i+1}: {desc_name} ({col})")
                    else:
                        print(f"  {i+1}: {desc_name}")
            except ImportError:
                # Fall back to original behavior if function not available
                print("\n  Available columns:")
                for i, col in enumerate(df.columns):
                    print(f"  {i+1}: {col}")
            
            # Ask for column selections
            while True:
                selection = input("\n  Enter column numbers to extract (comma-separated, 'all' for all, or 'done' to finish): ")
                
                if selection.lower() == 'done':
                    break
                elif selection.lower() == 'all':
                    selected_columns[file_name][sheet_name] = df.columns.tolist()
                    print(f"  Selected all {len(df.columns)} columns")
                    break
                else:
                    try:
                        # Parse the comma-separated input
                        col_indices = [int(idx.strip()) - 1 for idx in selection.split(',')]
                        
                        # Check if indices are valid
                        if any(idx < 0 or idx >= len(df.columns) for idx in col_indices):
                            print("  Error: Some column numbers are out of range. Please try again.")
                            continue
                        
                        # Add to selected columns
                        for idx in col_indices:
                            col = df.columns[idx]
                            if col not in selected_columns[file_name][sheet_name]:
                                selected_columns[file_name][sheet_name].append(col)
                        
                        print(f"  Selected: {', '.join(selected_columns[file_name][sheet_name])}")
                        
                        # Ask if they want to add more
                        add_more = input("  Add more columns from this sheet? (y/n): ")
                        if add_more.lower() != 'y':
                            break
                    
                    except Exception as e:
                        print(f"  Error: {e}. Please enter valid numbers separated by commas.")
            
            # Confirm selections for this sheet
            print(f"\n  Selected {len(selected_columns[file_name][sheet_name])} columns from {sheet_name}")
    
    # Provide a summary
    total_files = len(selected_columns)
    total_sheets = sum(1 for file in selected_columns.values() for sheet in file)
    total_columns = sum(len(cols) for file in selected_columns.values() for cols in file.values())
    
    print("\n=== SELECTION SUMMARY ===")
    print(f"Total files: {total_files}")
    print(f"Total sheets: {total_sheets}")
    print(f"Total columns selected: {total_columns}")
    
    return selected_columns

def main():
    # Set up argument parser
    parser = argparse.ArgumentParser(description='Extract and merge data from Excel files in a ZIP archive')
    parser.add_argument('zip_file', help='Path to the ZIP file containing Excel files')
    parser.add_argument('output_file', help='Path to save the merged Excel file')
    
    # Parse arguments
    if len(sys.argv) == 1:
        parser.print_help()
        sys.exit(1)
    
    args = parser.parse_args()
    
    # Get the ZIP file path
    zip_path = args.zip_file
    output_path = args.output_file
    
    # Check if files exist
    if not os.path.exists(zip_path):
        print(f"Error: ZIP file not found at {zip_path}")
        sys.exit(1)
    
    # Add extension if not present
    if not output_path.lower().endswith('.xls'):
        output_path += '.xls'
    
    # Create a temporary directory for extraction
    temp_dir = tempfile.mkdtemp()
    
    try:
        print("\n=== Excel Data Extractor CLI ===")
        print(f"Processing ZIP file: {zip_path}")
        
        # Extract Excel files from the ZIP
        print("\n=== EXTRACTING FILES ===")
        extracted_files = extract_zip_file(zip_path, temp_dir, print)
        
        if not extracted_files:
            print("No Excel files found in the ZIP archive.")
            sys.exit(1)
        
        # Read Excel files and store their data
        print("\n=== READING EXCEL FILES ===")
        file_data = read_excel_files(extracted_files, print)
        
        if not file_data:
            print("Could not read any data from the Excel files.")
            sys.exit(1)
        
        # Select columns to extract
        selected_columns = interactive_column_selection(file_data)
        
        # Check if any columns were selected
        total_selected = sum(len(cols) for file in selected_columns.values() for cols in file.values())
        if total_selected == 0:
            print("No columns were selected. Exiting.")
            sys.exit(1)
        
        # Process and merge the selected data
        print("\n=== PROCESSING AND MERGING DATA ===")
        output_dir = os.path.dirname(output_path)
        
        # Create output directory if it doesn't exist
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # Process and generate the output file
        success = process_and_merge_data(file_data, selected_columns, output_path, print)
        
        if success:
            print(f"\n=== PROCESSING COMPLETE ===")
            print(f"The merged Excel file has been saved to: {output_path}")
        else:
            print("\n=== PROCESSING FAILED ===")
            print("Failed to process and merge data.")
            sys.exit(1)
        
    finally:
        # Clean up temporary directory
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)

if __name__ == "__main__":
    main()