import streamlit as st
import os
import tempfile
import pandas as pd
from zipfile import ZipFile
import file_processor
import platform

# Set page configuration
st.set_page_config(
    page_title="Excel Data Extractor",
    page_icon="📊",
    layout="wide"
)

# Check if running on MacOS
is_macos = platform.system() == "Darwin"

# Initialize session state variables if they don't exist
if 'file_data' not in st.session_state:
    st.session_state.file_data = {}
if 'selected_columns' not in st.session_state:
    st.session_state.selected_columns = {}
if 'extracted_data' not in st.session_state:
    st.session_state.extracted_data = {}
if 'processing_stage' not in st.session_state:
    st.session_state.processing_stage = "upload"  # Stages: upload, selection, processing, complete
if 'output_path' not in st.session_state:
    st.session_state.output_path = None
if 'temp_dir' not in st.session_state:
    st.session_state.temp_dir = None
if 'zip_file_path' not in st.session_state:
    st.session_state.zip_file_path = None
if 'error_message' not in st.session_state:
    st.session_state.error_message = None
if 'debug_mode' not in st.session_state:
    st.session_state.debug_mode = False

def reset_app():
    # Clean up temporary files if they exist
    if st.session_state.temp_dir and os.path.exists(st.session_state.temp_dir):
        import shutil
        try:
            shutil.rmtree(st.session_state.temp_dir)
        except Exception as e:
            st.error(f"Error cleaning up temporary files: {e}")
    
    # Reset all session state variables
    st.session_state.file_data = {}
    st.session_state.selected_columns = {}
    st.session_state.extracted_data = {}
    st.session_state.processing_stage = "upload"
    st.session_state.output_path = None
    st.session_state.temp_dir = None
    st.session_state.zip_file_path = None
    st.session_state.error_message = None
    st.rerun()

# Title and description
st.title("Excel Data Extractor")
st.write("Upload a zip file containing Excel files, select data columns to extract, and merge them into a new Excel file.")

# Add MacOS-specific information
if is_macos:
    st.success("Running on MacOS - optimized for your system")
else:
    st.warning("This application is optimized for MacOS. Some features may not work correctly on other operating systems.")

# Add a debug mode toggle
with st.expander("Debug Options"):
    debug_enabled = st.checkbox("Enable Debug Mode", value=st.session_state.debug_mode)
    if debug_enabled != st.session_state.debug_mode:
        st.session_state.debug_mode = debug_enabled
        st.rerun()

# Show error message if any
if st.session_state.error_message:
    st.error(st.session_state.error_message)
    st.session_state.error_message = None

# Upload and extract zip file stage
if st.session_state.processing_stage == "upload":
    st.header("Step 1: Upload Zip File")
    
    # Guidance for MacOS users
    st.info("""
    **MacOS Tips:**
    1. Create a ZIP file by selecting multiple Excel files, right-clicking, and choosing "Compress"
    2. Make sure your Excel files are .xlsx format (Office Open XML)
    3. Avoid using special characters in filenames
    """)
    
    uploaded_file = st.file_uploader("Choose a ZIP file containing Excel (.xlsx) files", type=["zip"])
    
    if uploaded_file is not None:
        try:
            # Create a temporary directory to extract files
            temp_dir = tempfile.mkdtemp()
            st.session_state.temp_dir = temp_dir
            
            # Save the uploaded file to the temporary directory
            zip_path = os.path.join(temp_dir, "uploaded.zip")
            with open(zip_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            st.session_state.zip_file_path = zip_path
            
            # Show additional debug info if debug mode is enabled
            if st.session_state.debug_mode:
                st.write(f"ZIP file saved to: {zip_path}")
                st.write(f"Temporary directory: {temp_dir}")
            
            # Extract Excel files from the ZIP
            with st.spinner("Extracting ZIP file..."):
                extracted_files = file_processor.extract_zip_file(zip_path, temp_dir)
                
                if not extracted_files:
                    st.warning("No Excel files found in the ZIP archive.")
                    
                    # Show helpful troubleshooting info
                    with st.expander("Troubleshooting Tips"):
                        st.markdown("""
                        **Why might Excel files not be detected?**
                        - The ZIP file might not contain Excel files with .xlsx extension
                        - Excel files might be in subdirectories (we'll search for them)
                        - Files might be corrupted or password-protected
                        
                        **Try the following:**
                        - Enable Debug Mode (in the Debug Options expander above)
                        - Check that your Excel files have .xlsx extension
                        - Re-create the ZIP file on MacOS using Finder's "Compress" option
                        """)
                else:
                    # Read Excel files and store their data
                    st.session_state.file_data = file_processor.read_excel_files(extracted_files)
                    
                    if not st.session_state.file_data:
                        st.warning("Could not read any data from the Excel files.")
                        
                        # Show helpful troubleshooting info
                        with st.expander("Troubleshooting Tips"):
                            st.markdown("""
                            **Why might data not be readable?**
                            - Excel files might be empty
                            - Files might be corrupted or password-protected
                            - Files might have unusual formatting
                            
                            **Try the following:**
                            - Enable Debug Mode (in the Debug Options expander above)
                            - Check that your Excel files can be opened normally in Excel
                            - Save files in a simple Excel format without advanced features
                            """)
                    else:
                        # Initialize the selected_columns dictionary with the same structure as file_data
                        st.session_state.selected_columns = {}
                        for file_name, sheets in st.session_state.file_data.items():
                            st.session_state.selected_columns[file_name] = {}
                            for sheet_name, _ in sheets.items():
                                st.session_state.selected_columns[file_name][sheet_name] = []
                        
                        # Move to the next stage
                        st.session_state.processing_stage = "selection"
                        st.rerun()
        
        except Exception as e:
            st.session_state.error_message = f"Error processing ZIP file: {str(e)}"
            
            # Show more detailed error in debug mode
            if st.session_state.debug_mode:
                import traceback
                st.error(f"Error details:\n{traceback.format_exc()}")
            
            st.rerun()

# Data selection stage
elif st.session_state.processing_stage == "selection":
    st.header("Step 2: Select Data to Extract")
    
    # Display a counter to show how many files we're working with
    file_count = len(st.session_state.file_data.keys())
    st.write(f"Found {file_count} Excel file{'s' if file_count > 1 else ''} in the ZIP archive.")
    
    # Create tabs for each Excel file
    file_tabs = st.tabs(list(st.session_state.file_data.keys()))
    
    for i, (file_name, file_tab) in enumerate(zip(st.session_state.file_data.keys(), file_tabs)):
        with file_tab:
            st.subheader(f"File: {file_name}")
            
            # Create tabs for each sheet in the Excel file
            sheet_names = list(st.session_state.file_data[file_name].keys())
            if not sheet_names:
                st.warning(f"No sheets found in {file_name}")
                continue
                
            sheet_tabs = st.tabs(sheet_names)
            
            for j, (sheet_name, sheet_tab) in enumerate(zip(sheet_names, sheet_tabs)):
                with sheet_tab:
                    st.write(f"Sheet: {sheet_name}")
                    
                    # Get the dataframe for this sheet
                    df = st.session_state.file_data[file_name][sheet_name]
                    
                    # Display a preview of the data
                    st.write("Data Preview:")
                    st.dataframe(df.head(5), use_container_width=True)
                    
                    # Create checkboxes for column selection
                    st.write("Select columns to extract:")
                    
                    # Use st.columns to create a grid of checkboxes
                    cols_per_row = 3
                    all_columns = df.columns.tolist()
                    
                    # Select all / Deselect all buttons
                    col1, col2 = st.columns(2)
                    with col1:
                        if st.button(f"Select All ({sheet_name})", key=f"select_all_{file_name}_{sheet_name}"):
                            st.session_state.selected_columns[file_name][sheet_name] = all_columns
                            st.rerun()
                    with col2:
                        if st.button(f"Deselect All ({sheet_name})", key=f"deselect_all_{file_name}_{sheet_name}"):
                            st.session_state.selected_columns[file_name][sheet_name] = []
                            st.rerun()
                    
                    # Display checkboxes in a grid layout
                    for k in range(0, len(all_columns), cols_per_row):
                        cols = st.columns(cols_per_row)
                        for l in range(cols_per_row):
                            idx = k + l
                            if idx < len(all_columns):
                                col_name = all_columns[idx]
                                # Check if this column is already selected
                                is_selected = col_name in st.session_state.selected_columns[file_name][sheet_name]
                                
                                # Create a unique key for each checkbox
                                checkbox_key = f"cb_{file_name}_{sheet_name}_{col_name}"
                                
                                # Create the checkbox
                                with cols[l]:
                                    if st.checkbox(col_name, value=is_selected, key=checkbox_key):
                                        if col_name not in st.session_state.selected_columns[file_name][sheet_name]:
                                            st.session_state.selected_columns[file_name][sheet_name].append(col_name)
                                    else:
                                        if col_name in st.session_state.selected_columns[file_name][sheet_name]:
                                            st.session_state.selected_columns[file_name][sheet_name].remove(col_name)
    
    # Show how many columns are selected in total
    total_selected = sum(len(cols) for file in st.session_state.selected_columns.values() 
                         for cols in file.values())
    
    st.write(f"Total columns selected: {total_selected}")
    
    # Provide buttons for proceeding or going back
    col1, col2 = st.columns(2)
    with col1:
        if st.button("Back to Upload", use_container_width=True):
            reset_app()
    
    with col2:
        if st.button("Continue to Processing", use_container_width=True):
            # Validate that at least one column is selected
            if total_selected == 0:
                st.warning("Please select at least one column to extract.")
            else:
                # Get output folder path (actually we'll get output filename in MacOS)
                st.session_state.processing_stage = "processing"
                st.rerun()

# Output selection and processing stage
elif st.session_state.processing_stage == "processing":
    st.header("Step 3: Select Output File and Process Data")
    
    # Prompt for output file path
    st.write("Please enter a name for the output Excel file:")
    output_filename = st.text_input("Output filename (without extension)", "merged_data")
    
    # Add extension if not present
    if not output_filename.endswith(".xls"):
        output_filename += ".xls"
    
    if st.button("Process and Generate Output File"):
        try:
            with st.spinner("Processing data..."):
                # Create a temporary output file
                temp_output_path = os.path.join(st.session_state.temp_dir, output_filename)
                
                # Process the selected data
                file_processor.process_and_merge_data(
                    st.session_state.file_data,
                    st.session_state.selected_columns,
                    temp_output_path
                )
                
                # Create a download link for the generated file
                with open(temp_output_path, "rb") as f:
                    file_bytes = f.read()
                    st.download_button(
                        label="Download Merged Excel File",
                        data=file_bytes,
                        file_name=output_filename,
                        mime="application/vnd.ms-excel"
                    )
                
                st.success(f"Data successfully extracted and merged into {output_filename}")
                st.session_state.processing_stage = "complete"
                st.rerun()
        
        except Exception as e:
            st.error(f"Error processing data: {str(e)}")
    
    if st.button("Back to Selection"):
        st.session_state.processing_stage = "selection"
        st.rerun()

# Final stage - show results and offer to start over
elif st.session_state.processing_stage == "complete":
    st.header("Processing Complete!")
    st.success("Your data has been successfully extracted and merged.")
    
    # Show a summary of what was processed
    st.subheader("Processing Summary")
    st.write(f"Number of files processed: {len(st.session_state.file_data)}")
    
    total_sheets = sum(len(sheets) for sheets in st.session_state.file_data.values())
    st.write(f"Number of sheets processed: {total_sheets}")
    
    total_cols = sum(len(cols) for file in st.session_state.selected_columns.values() 
                    for cols in file.values())
    st.write(f"Number of columns extracted: {total_cols}")
    
    # Offer to download the file again
    temp_output_path = os.path.join(st.session_state.temp_dir, [f for f in os.listdir(st.session_state.temp_dir) if f.endswith('.xls')][0])
    with open(temp_output_path, "rb") as f:
        file_bytes = f.read()
        st.download_button(
            label="Download Merged Excel File Again",
            data=file_bytes,
            file_name=os.path.basename(temp_output_path),
            mime="application/vnd.ms-excel"
        )
    
    # Button to start over
    if st.button("Start Over", use_container_width=True):
        reset_app()

# Add a footer
st.markdown("---")
st.caption("Excel Data Extractor | Optimized for MacOS")
