import os
import pandas as pd
from zipfile import ZipFile
import tempfile
import streamlit as st
from pathlib import Path
import xlwt

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
        with ZipFile(zip_path, 'r') as zip_ref:
            # List all files in the ZIP
            file_list = zip_ref.namelist()
            
            # Extract only Excel files
            for file_name in file_list:
                if file_name.lower().endswith('.xlsx'):
                    # Extract the file
                    zip_ref.extract(file_name, extract_dir)
                    excel_files.append(os.path.join(extract_dir, file_name))
    except Exception as e:
        st.error(f"Error extracting ZIP file: {str(e)}")
        return []
    
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
    
    for file_path in file_paths:
        try:
            # Get just the filename without path
            file_name = os.path.basename(file_path)
            
            # Read all sheets from the Excel file
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names
            
            # Initialize the entry for this file
            file_data[file_name] = {}
            
            # Read each sheet and store its data
            for sheet_name in sheet_names:
                try:
                    df = pd.read_excel(excel_file, sheet_name=sheet_name)
                    
                    # Only keep sheets that have data
                    if not df.empty:
                        file_data[file_name][sheet_name] = df
                except Exception as e:
                    st.warning(f"Error reading sheet '{sheet_name}' in file '{file_name}': {str(e)}")
                    continue
            
            # If no sheets were successfully read, remove this file entry
            if not file_data[file_name]:
                del file_data[file_name]
                
        except Exception as e:
            st.warning(f"Error reading file '{os.path.basename(file_path)}': {str(e)}")
            continue
    
    return file_data

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
        # Create a new workbook
        workbook = xlwt.Workbook()
        
        # Track the number of worksheets created
        worksheet_count = 0
        
        # Process each file
        for file_name, sheets in file_data.items():
            # Process each sheet in the file
            for sheet_name, df in sheets.items():
                # Get the selected columns for this sheet
                cols = selected_columns.get(file_name, {}).get(sheet_name, [])
                
                # Skip if no columns were selected for this sheet
                if not cols:
                    continue
                
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
        workbook.save(output_path)
        
        return True
    
    except Exception as e:
        st.error(f"Error processing and merging data: {str(e)}")
        raise e
