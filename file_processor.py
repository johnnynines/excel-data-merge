#!/usr/bin/env python3
"""
Excel Data Extractor - Core Processing Module
This module contains shared functionality for processing Excel files in a ZIP archive.
"""

import os
import pandas as pd
from zipfile import ZipFile
import xlwt
import tempfile
import re

def extract_zip_file(zip_path, extract_dir, log_callback=None):
    """
    Extract Excel files from a ZIP archive
    
    Parameters:
    - zip_path: Path to the ZIP file
    - extract_dir: Directory to extract files to
    - log_callback: Optional callback function for logging
    
    Returns:
    - A list of paths to extracted Excel files
    """
    excel_files = []
    
    try:
        if log_callback:
            log_callback(f"Opening ZIP file: {zip_path}")
        
        with ZipFile(zip_path, 'r') as zip_ref:
            # List all files in the ZIP
            file_list = zip_ref.namelist()
            
            if log_callback:
                log_callback(f"Found {len(file_list)} files in ZIP archive")
            
            # Extract only Excel files
            for file_name in file_list:
                lower_name = file_name.lower()
                if lower_name.endswith('.xlsx') or lower_name.endswith('.xls'):
                    # Handle folder paths in ZIP
                    if file_name.endswith('/') or os.path.basename(file_name) == '':
                        continue
                        
                    # Extract the file
                    try:
                        if log_callback:
                            log_callback(f"Extracting: {file_name}")
                        zip_ref.extract(file_name, extract_dir)
                        full_path = os.path.join(extract_dir, file_name)
                        excel_files.append(full_path)
                    except Exception as extract_error:
                        if log_callback:
                            log_callback(f"Could not extract {file_name}: {str(extract_error)}")
                
            # Also look for Excel files in any folders that were extracted
            for root, dirs, files in os.walk(extract_dir):
                for file in files:
                    if file.lower().endswith(('.xlsx', '.xls')) and os.path.join(root, file) not in excel_files:
                        excel_files.append(os.path.join(root, file))
                        if log_callback:
                            log_callback(f"Found additional Excel file: {file}")
    
    except Exception as e:
        if log_callback:
            log_callback(f"Error extracting ZIP file: {str(e)}")
        return []
    
    if log_callback:
        log_callback(f"Extracted {len(excel_files)} Excel files")
    return excel_files

def read_excel_files(file_paths, log_callback=None):
    """
    Read data from multiple Excel files
    
    Parameters:
    - file_paths: List of paths to Excel files
    - log_callback: Optional callback function for logging
    
    Returns:
    - A nested dictionary structure: {file_name: {sheet_name: dataframe}}
    """
    file_data = {}
    
    if not file_paths:
        if log_callback:
            log_callback("No Excel files to process")
        return file_data
    
    if log_callback:
        log_callback(f"Reading {len(file_paths)} Excel files...")
    
    for file_path in file_paths:
        try:
            # Get just the filename without path
            file_name = os.path.basename(file_path)
            if log_callback:
                log_callback(f"Reading: {file_name}")
            
            # Read all sheets from the Excel file
            try:
                excel_file = pd.ExcelFile(file_path)
                sheet_names = excel_file.sheet_names
                if log_callback:
                    log_callback(f"Found {len(sheet_names)} sheets in {file_name}")
            except Exception as excel_error:
                if log_callback:
                    log_callback(f"Error opening Excel file '{file_name}': {str(excel_error)}")
                
                # Try alternate approach for older Excel formats
                try:
                    # For xls files
                    if file_path.lower().endswith('.xls'):
                        df = pd.read_excel(file_path, engine='xlrd')
                        file_data[file_name] = {"Sheet1": df}
                        if log_callback:
                            log_callback(f"Successfully read {file_name} using xlrd engine")
                        continue
                except Exception as alt_error:
                    if log_callback:
                        log_callback(f"Alternative read approach failed: {str(alt_error)}")
                continue
            
            # Initialize the entry for this file
            file_data[file_name] = {}
            
            # Read each sheet and store its data
            for sheet_name in sheet_names:
                try:
                    # IMPROVED APPROACH: Intelligently detect column headers
                    # First grab the raw data without assuming header position
                    raw_df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
                    
                    if log_callback:
                        log_callback(f"Raw sheet '{sheet_name}' has {len(raw_df)} rows and {len(raw_df.columns)} columns")
                    
                    # If dataframe is completely empty, skip it
                    if raw_df.empty:
                        if log_callback:
                            log_callback(f"Sheet '{sheet_name}' is completely empty, skipping")
                        continue
                    
                    # Detect header row by checking for non-empty rows
                    header_row = 0
                    max_check_rows = min(10, len(raw_df))  # Look at most in the first 10 rows
                    
                    # Look for the first non-empty row to use as headers
                    for i in range(max_check_rows):
                        # Check if this row has mostly non-null values
                        row_values = raw_df.iloc[i].dropna()
                        if len(row_values) > 0 and len(row_values) >= len(raw_df.columns) / 2:
                            header_row = i
                            if log_callback:
                                log_callback(f"Found potential header row at index {header_row}")
                            break
                    
                    # Extract headers from the detected row
                    if header_row > 0:
                        if log_callback:
                            log_callback(f"Using row {header_row+1} as header instead of first row")
                        headers = raw_df.iloc[header_row].tolist()
                        # Clean up headers - convert to strings and replace NaN with generic names
                        headers = [f"Column_{i}" if pd.isna(h) else str(h).strip() for i, h in enumerate(headers)]
                        
                        # Create a dataframe with these headers, skipping the header row
                        data_rows = list(range(0, header_row)) + list(range(header_row+1, len(raw_df)))
                        df = pd.DataFrame(raw_df.iloc[data_rows].values, columns=headers)
                        
                        if log_callback:
                            header_sample = ', '.join(headers[:min(5, len(headers))])
                            if len(headers) > 5:
                                header_sample += "..."
                            log_callback(f"Found headers: {header_sample}")
                    else:
                        # No suitable header row found - use generic column names
                        if log_callback:
                            log_callback(f"Using generic column names (no clear header row found)")
                        column_names = [f"Column_{i}" for i in range(len(raw_df.columns))]
                        df = pd.DataFrame(raw_df.values, columns=column_names)
                    
                    # Store this dataframe even if it has blank rows - important to not lose data
                    file_data[file_name][sheet_name] = df
                    
                    if log_callback:
                        log_callback(f"Successfully processed sheet '{sheet_name}' with {len(df)} rows and {len(df.columns)} columns")
                except Exception as e:
                    if log_callback:
                        log_callback(f"Error reading sheet '{sheet_name}': {str(e)}")
                    continue
            
            # If no sheets were successfully read, remove this file entry
            if not file_data[file_name]:
                if log_callback:
                    log_callback(f"No data found in file '{file_name}'")
                del file_data[file_name]
                
        except Exception as e:
            if log_callback:
                log_callback(f"Error reading file '{os.path.basename(file_path)}': {str(e)}")
            continue
    
    # Provide summary
    file_count = len(file_data)
    if file_count > 0:
        sheet_count = sum(len(sheets) for sheets in file_data.values())
        if log_callback:
            log_callback(f"Successfully read {file_count} files with a total of {sheet_count} sheets")
    else:
        if log_callback:
            log_callback("Could not read any data from the Excel files")
    
    return file_data

def detect_descriptive_column_names(df, log_callback=None):
    """
    Detects more descriptive column names by finding the first non-empty string value in each column.
    Used to provide more meaningful headers in the UI.
    
    Parameters:
    - df: pandas DataFrame to analyze
    - log_callback: Optional callback function for logging
    
    Returns:
    - Dictionary mapping original column names to detected descriptive names
    """
    descriptive_names = {}
    
    # Skip if DataFrame is empty
    if df is None or df.empty:
        return descriptive_names
        
    # Process each column to find the first non-empty string
    for col in df.columns:
        # Default to the original column name (ensure it's a string)
        col_str = str(col)
        descriptive_names[col] = col_str
        
        try:
            # Get non-empty string values from this column (but only consider the first 20 rows for efficiency)
            sample_size = min(20, len(df))
            sample = df[col].head(sample_size)
            
            # Look for the first non-empty string value that is not just a number
            for value in sample:
                if pd.notna(value) and isinstance(value, str) and value.strip() and not value.strip().isdigit():
                    # Clean up the value to use as a header (max 30 chars to stay readable)
                    desc_name = str(value).strip()
                    # Truncate if too long, but preserve meaningful content
                    if len(desc_name) > 30:
                        desc_name = desc_name[:27] + "..."
                    
                    # Only use it if it's better than a generic column name
                    is_generic = isinstance(col_str, str) and col_str.startswith("Column_")
                    if not is_generic or len(desc_name) > 0:
                        descriptive_names[col] = desc_name
                    break
        except Exception as e:
            if log_callback:
                log_callback(f"Error detecting descriptive name for column {col}: {str(e)}")
                
    if log_callback:
        log_callback(f"Detected {len(descriptive_names)} descriptive column names")
        
    return descriptive_names

def process_and_merge_data(file_data, selected_columns, output_path, log_callback=None):
    """
    Process and merge selected data from multiple Excel files
    
    Parameters:
    - file_data: Nested dictionary of file data {file_name: {sheet_name: dataframe}}
    - selected_columns: Nested dictionary of selected columns {file_name: {sheet_name: [columns]}}
    - output_path: Path to save the merged Excel file
    - log_callback: Optional callback function for logging
    
    Returns:
    - True if successful, False otherwise
    """
    try:
        if log_callback:
            log_callback("Starting data processing...")
        
        # Create a new workbook
        workbook = xlwt.Workbook()
        
        # Track the number of worksheets created
        worksheet_count = 0
        
        # Process each file
        for file_name, sheets in file_data.items():
            if log_callback:
                log_callback(f"Processing file: {file_name}")
            
            # Process each sheet in the file
            for sheet_name, df in sheets.items():
                # Get the selected columns for this sheet
                cols = selected_columns.get(file_name, {}).get(sheet_name, [])
                
                # Skip if no columns were selected for this sheet
                if not cols:
                    if log_callback:
                        log_callback(f"No columns selected for {file_name} - {sheet_name}, skipping")
                    continue
                
                if log_callback:
                    log_callback(f"Processing sheet: {sheet_name} with {len(cols)} selected columns")
                
                # Extract only the selected columns
                subset_df = df[cols].copy()
                
                # Create a worksheet name from the file and sheet names
                # Ensure it's valid and not too long for Excel
                from pathlib import Path
                ws_name = f"{Path(file_name).stem}_{sheet_name}"
                ws_name = ws_name.replace("[", "").replace("]", "").replace(":", "")
                ws_name = ws_name[:31]  # Excel has 31 char limit for sheet names
                
                # Handle duplicate sheet names by appending a number
                original_ws_name = ws_name
                counter = 1
                # Get existing worksheet names - xlwt doesn't have get_sheets() method
                existing_sheet_names = [sheet.name for sheet in workbook._Workbook__worksheets]
                while ws_name in existing_sheet_names:
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
                    # Convert all column names to strings before joining
                    cols_str = [str(col) for col in cols]
                    summary.write(row, 2, ", ".join(cols_str))
                    row += 1
        
        # Save the workbook
        if log_callback:
            log_callback(f"Saving output to: {output_path}")
        workbook.save(output_path)
        
        if log_callback:
            log_callback(f"Processing complete. Created {worksheet_count} worksheets plus summary.")
        return True
    
    except Exception as e:
        if log_callback:
            log_callback(f"Error processing and merging data: {str(e)}")
        return False