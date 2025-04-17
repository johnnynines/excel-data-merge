#!/usr/bin/env python3
"""
Excel Data Extractor - PyQt5 macOS Application
This application extracts and merges selected data from multiple Excel files in a ZIP archive.
Optimized specifically for macOS with native look and feel.
"""

import sys
import os
import tempfile
import pandas as pd
import zipfile
from pathlib import Path
import xlwt

# Set environment variable for improved macOS look and feel
os.environ['QT_MAC_WANTS_LAYER'] = '1'  # Improves rendering on macOS

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QMessageBox, QProgressBar,
    QTabWidget, QCheckBox, QGroupBox, QScrollArea, QGridLayout,
    QLineEdit, QTableView, QHeaderView, QSplitter, QFrame, QStyle,
    QTreeWidget, QTreeWidgetItem, QStackedWidget
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QAbstractTableModel, QModelIndex, QSize
from PyQt5.QtGui import QFont, QIcon, QPalette, QColor

# Model for displaying Excel data in a table
class PandasTableModel(QAbstractTableModel):
    def __init__(self, data):
        """
        Initialize the table model with pandas DataFrame data
        
        Parameters:
        - data: pandas DataFrame or something that can be converted to one
        """
        super().__init__()
        
        # Ensure we're working with a DataFrame
        if not isinstance(data, pd.DataFrame):
            try:
                data = pd.DataFrame(data)
            except:
                # Create an empty dataframe with a message if data is invalid
                data = pd.DataFrame({'Error': ['Invalid data provided to table model']})
        
        # Store original data
        self._original_data = data
        
        # Process the dataframe for display
        self._process_dataframe()
    
    def _process_dataframe(self):
        """
        Clean and prepare the dataframe for display
        SIMPLIFIED: No blank row handling, just make it displayable
        """
        # Make a copy to avoid modifying original
        self._data = self._original_data.copy()
        
        # Handle completely empty dataframes
        if self._data.empty:
            self._data = pd.DataFrame({'No Data': ['Empty sheet - no data to display']})
            return
            
        # Ensure column names are strings (this is required for display)
        self._data.columns = [str(col) if not pd.isna(col) else f"Column_{i}" 
                             for i, col in enumerate(self._data.columns)]

    def rowCount(self, parent=None):
        """Return the number of rows in the dataframe"""
        if parent and parent.isValid():
            return 0
        return len(self._data)

    def columnCount(self, parent=None):
        """Return the number of columns in the dataframe"""
        if parent and parent.isValid():
            return 0
        return len(self._data.columns)

    def data(self, index, role=Qt.DisplayRole):
        """Return the data at the given index for the specified role"""
        if not index.isValid():
            return None
            
        if role == Qt.DisplayRole or role == Qt.EditRole:
            try:
                value = self._data.iloc[index.row(), index.column()]
                # Handle NaN and None values properly
                if pd.isna(value):
                    return ""
                # Convert all values to string for display
                return str(value)
            except (IndexError, KeyError):
                return ""
        
        # Add styling for alternate rows
        if role == Qt.BackgroundRole:
            if index.row() % 2 == 0:
                # Light background for even rows
                return QColor(248, 248, 248)
        
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        """Return the header data for the specified section, orientation and role"""
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                try:
                    # Use column name for horizontal headers
                    return str(self._data.columns[section])
                except IndexError:
                    # Fallback to section number
                    return f"Column_{section}"
            else:
                # Row numbers for vertical header (1-based)
                return str(section + 1)
        
        # Add styling for headers
        if role == Qt.FontRole:
            font = QFont()
            font.setBold(True)
            return font
            
        return None

# Worker thread for processing files
class FileProcessorThread(QThread):
    progress_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(dict)
    error_signal = pyqtSignal(str)
    
    def __init__(self, zip_path, extract_dir):
        super().__init__()
        self.zip_path = zip_path
        self.extract_dir = extract_dir
        
    def run(self):
        try:
            # Extract Excel files from ZIP
            excel_files = self.extract_zip_file()
            
            if not excel_files:
                self.error_signal.emit("No Excel files found in the ZIP archive.")
                return
                
            # Read Excel files
            file_data = self.read_excel_files(excel_files)
            
            if not file_data:
                self.error_signal.emit("Could not read any data from the Excel files.")
                return
                
            # Signal completion
            self.finished_signal.emit(file_data)
            
        except Exception as e:
            self.error_signal.emit(f"Error processing files: {str(e)}")
    
    def extract_zip_file(self):
        """Extract Excel files from a ZIP archive"""
        excel_files = []
        found_files = set()  # Use a set to track unique files by name
        
        try:
            self.progress_signal.emit("Opening ZIP file...")
            
            with zipfile.ZipFile(self.zip_path, 'r') as zip_ref:
                # List all files in the ZIP
                file_list = zip_ref.namelist()
                
                self.progress_signal.emit(f"Found {len(file_list)} files in ZIP archive")
                
                # Debug: Print all files in the ZIP for inspection
                all_files_str = ", ".join([f for f in file_list if not f.endswith('/')])
                self.progress_signal.emit(f"Files in ZIP: {all_files_str}")
                
                # Extract only Excel files
                for file_name in file_list:
                    lower_name = file_name.lower()
                    if lower_name.endswith('.xlsx') or lower_name.endswith('.xls'):
                        # Handle folder paths in ZIP
                        if file_name.endswith('/') or os.path.basename(file_name) == '':
                            continue
                        
                        # Extract the file
                        try:
                            # Log the exact file name for debugging
                            base_name = os.path.basename(file_name)
                            self.progress_signal.emit(f"Extracting Excel file: {base_name} (full path: {file_name})")
                            
                            # Extract the file
                            zip_ref.extract(file_name, self.extract_dir)
                            full_path = os.path.join(self.extract_dir, file_name)
                            
                            # Make sure we don't add duplicates
                            if full_path not in excel_files:
                                excel_files.append(full_path)
                                found_files.add(base_name)
                                self.progress_signal.emit(f"Added to processing list: {base_name}")
                        except Exception as extract_error:
                            self.progress_signal.emit(f"Could not extract {file_name}: {str(extract_error)}")
                    
                # Also look for Excel files in extracted folders that might have been missed
                skipped_files = []
                for root, dirs, files in os.walk(self.extract_dir):
                    for file in files:
                        if file.lower().endswith(('.xlsx', '.xls')):
                            full_path = os.path.join(root, file)
                            if full_path not in excel_files:
                                excel_files.append(full_path)
                                if file not in found_files:
                                    found_files.add(file)
                                    self.progress_signal.emit(f"Found additional Excel file: {file}")
                                else:
                                    # File was found but under a different path
                                    self.progress_signal.emit(f"NOTE: Found duplicate Excel file with different path: {file}")
                            else:
                                skipped_files.append(file)
                
                if skipped_files:
                    self.progress_signal.emit(f"Skipped duplicates: {', '.join(skipped_files)}")
                
                # Debug - list all extracted files
                self.progress_signal.emit(f"All extracted Excel files: {', '.join(found_files)}")
        
        except Exception as e:
            self.error_signal.emit(f"Error extracting ZIP file: {str(e)}")
            return []
        
        # Make sure all Excel files are unique by path
        unique_files = []
        seen_paths = set()
        for file_path in excel_files:
            if file_path not in seen_paths:
                unique_files.append(file_path)
                seen_paths.add(file_path)
        
        self.progress_signal.emit(f"Found {len(unique_files)} unique Excel files. Processing now...")
        
        # Sort files alphabetically to ensure consistent processing order
        unique_files.sort()
        
        # Final verification
        self.progress_signal.emit(f"Files to be processed:")
        for i, file_path in enumerate(unique_files):
            self.progress_signal.emit(f"{i+1}. {os.path.basename(file_path)}")
            
        return unique_files
    
    def read_excel_files(self, file_paths):
        """Read data from multiple Excel files"""
        file_data = {}  # This will store our processed Excel data
        processed_files = set()  # Keep track of processed files to detect issues
        
        if not file_paths:
            self.progress_signal.emit("No Excel files to process")
            return file_data
        
        self.progress_signal.emit(f"Reading {len(file_paths)} Excel files...")
        
        # For debugging - explicitly list all files we'll process
        for idx, file_path in enumerate(file_paths):
            self.progress_signal.emit(f"Will process #{idx+1}: {os.path.basename(file_path)}")
        
        # Track the original file names to make sure we don't lose any
        original_filenames = [os.path.basename(path) for path in file_paths]
        self.progress_signal.emit(f"Original filenames to process: {', '.join(original_filenames)}")
        
        # Process each file in the list
        for file_idx, file_path in enumerate(file_paths):
            try:
                # Get just the filename without path
                raw_file_name = os.path.basename(file_path)
                
                # Check if file exists
                if not os.path.exists(file_path):
                    self.progress_signal.emit(f"ERROR: File does not exist: {file_path}")
                    continue
                
                # Better file name sanitization
                file_name = raw_file_name
                # Replace problematic characters with underscore
                for char in [' ', '-', '(', ')', '[', ']', '{', '}', '&', '+', '=']:
                    file_name = file_name.replace(char, '_')
                
                # Log both original and sanitized filenames
                self.progress_signal.emit(f"Processing file {file_idx+1}/{len(file_paths)}: {raw_file_name}")
                if raw_file_name != file_name:
                    self.progress_signal.emit(f"Using sanitized file name: {file_name} for internal processing")
                
                # Track this file as processed
                processed_files.add(raw_file_name)
                
                # Read all sheets from the Excel file
                try:
                    # Try pandas ExcelFile first
                    self.progress_signal.emit(f"Attempting to read: {file_path}")
                    excel_file = pd.ExcelFile(file_path)
                    sheet_names = excel_file.sheet_names
                    self.progress_signal.emit(f"Found {len(sheet_names)} sheets in {file_name}: {', '.join(sheet_names)}")
                except Exception as excel_error:
                    self.progress_signal.emit(f"Error opening Excel file '{file_name}': {str(excel_error)}")
                    
                    # Try alternate approach for older Excel formats
                    try:
                        # For xls files
                        if file_path.lower().endswith('.xls'):
                            self.progress_signal.emit(f"Trying alternate read method with xlrd engine")
                            df = pd.read_excel(file_path, engine='xlrd')
                            file_data[file_name] = {"Sheet1": df}
                            self.progress_signal.emit(f"Successfully read {file_name} using xlrd engine")
                            continue
                    except Exception as alt_error:
                        self.progress_signal.emit(f"Alternative read approach failed: {str(alt_error)}")
                    self.progress_signal.emit(f"SKIPPING file {file_name} due to errors")
                    continue
                
                # Print debugging info about file data before adding
                existing_files = list(file_data.keys())
                self.progress_signal.emit(f"Current files in data dictionary: {existing_files}")
                
                # Initialize the entry for this file, ensuring we don't overwrite existing data
                if file_name in file_data:
                    self.progress_signal.emit(f"WARNING: File with name {file_name} already exists in data! Adding unique suffix...")
                    base_name = file_name
                    counter = 1
                    while file_name in file_data:
                        file_name = f"{base_name}_{counter}"
                        counter += 1
                    self.progress_signal.emit(f"Using unique file name: {file_name}")
                
                # Create the dictionary entry for this file
                file_data[file_name] = {}
                
                # Read each sheet and store its data
                for sheet_name in sheet_names:
                    try:
                        # We've moved this logic to the shared file_processor.py module
                        # Let the shared functionality handle Excel processing
                        try:
                            from file_processor import read_excel_files
                            
                            # Create a temporary dict just for this file/sheet
                            temp_excel_file = {file_name: {}}
                            
                            # Read this sheet using the shared logic
                            self.progress_signal.emit(f"Using enhanced header detection for {sheet_name}")
                            
                            # Process the sheet through the shared processor
                            temp_excel_file[file_name][sheet_name] = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
                            
                            # Process this file with our improved shared header detection
                            result = read_excel_files([excel_file], lambda msg: self.progress_signal.emit(msg))
                            
                            # If processing was successful, copy the processed sheet
                            if file_name in result and sheet_name in result[file_name]:
                                df = result[file_name][sheet_name]
                                file_data[file_name][sheet_name] = df
                            else:
                                # Fallback if something went wrong
                                self.progress_signal.emit(f"Could not process sheet through enhanced detection, using basic process")
                                raw_df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
                                column_names = [f"Column_{i}" for i in range(len(raw_df.columns))]
                                df = pd.DataFrame(raw_df.values, columns=column_names)
                                file_data[file_name][sheet_name] = df
                                
                        except Exception as e:
                            # If there's any issue, fall back to basic processing
                            self.progress_signal.emit(f"Error using enhanced header detection: {str(e)}")
                            raw_df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
                            column_names = [f"Column_{i}" for i in range(len(raw_df.columns))]
                            df = pd.DataFrame(raw_df.values, columns=column_names)
                            file_data[file_name][sheet_name] = df
                        
                        self.progress_signal.emit(f"Successfully processed sheet '{sheet_name}' with {len(df)} rows and {len(df.columns)} columns")
                    except Exception as e:
                        self.progress_signal.emit(f"Error reading sheet '{sheet_name}': {str(e)}")
                        continue
                
                # If no sheets were successfully read, remove this file entry
                if not file_data[file_name]:
                    self.progress_signal.emit(f"No data found in file '{file_name}'")
                    del file_data[file_name]
                    
            except Exception as e:
                self.progress_signal.emit(f"Error reading file '{os.path.basename(file_path)}': {str(e)}")
                continue
        
        # Provide summary
        file_count = len(file_data)
        if file_count > 0:
            sheet_count = sum(len(sheets) for sheets in file_data.values())
            self.progress_signal.emit(f"Successfully read {file_count} files with {sheet_count} sheets")
        else:
            self.progress_signal.emit("Could not read any data from the Excel files")
        
        return file_data

# Worker thread for processing output
class OutputProcessorThread(QThread):
    progress_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(str)
    error_signal = pyqtSignal(str)
    
    def __init__(self, file_data, selected_columns, output_path):
        super().__init__()
        self.file_data = file_data
        self.selected_columns = selected_columns
        self.output_path = output_path
        
    def run(self):
        try:
            self.progress_signal.emit("Starting data processing...")
            self.process_and_merge_data()
            self.finished_signal.emit(self.output_path)
        except Exception as e:
            self.error_signal.emit(f"Error processing data: {str(e)}")
    
    def process_and_merge_data(self):
        """Process and merge selected data from multiple Excel files"""
        try:
            # Create a new workbook
            workbook = xlwt.Workbook()
            
            # Track the number of worksheets created
            worksheet_count = 0
            
            # Process each file
            for file_name, sheets in self.file_data.items():
                self.progress_signal.emit(f"Processing file: {file_name}")
                
                # Process each sheet in the file
                for sheet_name, df in sheets.items():
                    # Get the selected columns for this sheet
                    cols = self.selected_columns.get(file_name, {}).get(sheet_name, [])
                    
                    # Skip if no columns were selected for this sheet
                    if not cols:
                        self.progress_signal.emit(f"No columns selected for {file_name} - {sheet_name}, skipping")
                        continue
                    
                    self.progress_signal.emit(f"Processing sheet: {sheet_name} with {len(cols)} selected columns")
                    
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
            for file_name, sheets in self.selected_columns.items():
                for sheet_name, cols in sheets.items():
                    if cols:  # Only include sheets where columns were selected
                        summary.write(row, 0, file_name)
                        summary.write(row, 1, sheet_name)
                        summary.write(row, 2, ", ".join(cols))
                        row += 1
            
            # Save the workbook
            self.progress_signal.emit(f"Saving output to: {self.output_path}")
            workbook.save(self.output_path)
            
            self.progress_signal.emit(f"Processing complete. Created {worksheet_count} worksheets plus summary.")
            return True
        
        except Exception as e:
            self.error_signal.emit(f"Error processing and merging data: {str(e)}")
            raise e
            
class ExcelExtractorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # Initialize instance variables
        self.file_data = {}
        self.selected_columns = {}
        self.temp_dir = None
        self.output_path = None
        
        # Setup UI
        self.init_ui()
        
    def init_ui(self):
        # Set window properties with macOS optimizations
        self.setWindowTitle("Excel Data Extractor")
        self.setGeometry(100, 100, 900, 600)
        
        # Set application icon using system icon (document icon on macOS)
        self.setWindowIcon(self.style().standardIcon(QStyle.SP_FileDialogDetailedView))
        
        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # macOS specific styling - ensure proper spacing and margins
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(10)
        
        # Create header
        header_label = QLabel("Excel Data Extractor")
        header_label.setFont(QFont("Arial", 18, QFont.Bold))
        header_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(header_label)
        
        desc_label = QLabel("Extract and merge data from multiple Excel files in a ZIP archive")
        desc_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(desc_label)
        
        # Create tab widget for different stages
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)
        
        # Create tabs for each stage
        self.upload_tab = QWidget()
        self.selection_tab = QTabWidget()  # This will be a nested tab widget
        self.output_tab = QWidget()
        
        self.tabs.addTab(self.upload_tab, "1. Upload ZIP")
        self.tabs.addTab(self.selection_tab, "2. Select Data")
        self.tabs.addTab(self.output_tab, "3. Generate Output")
        
        # Disable tabs until they're ready
        self.tabs.setTabEnabled(1, False)
        self.tabs.setTabEnabled(2, False)
        
        # Setup each tab
        self.setup_upload_tab()
        self.setup_output_tab()
        
        # Status bar
        self.statusBar().showMessage("Ready")
        
        # Progress bar in status bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setMaximumWidth(300)
        self.progress_bar.setMaximumHeight(16)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setVisible(False)
        self.statusBar().addPermanentWidget(self.progress_bar)
        
    def setup_upload_tab(self):
        """Setup UI for the upload tab"""
        layout = QVBoxLayout(self.upload_tab)
        
        # Instructions
        instruction_label = QLabel(
            "Select a ZIP file containing Excel files (.xlsx or .xls).\n"
            "The application will extract the Excel files and let you choose which data to extract."
        )
        instruction_label.setWordWrap(True)
        layout.addWidget(instruction_label)
        
        # MacOS Tips
        tips_group = QGroupBox("MacOS Tips")
        tips_layout = QVBoxLayout()
        tips_label = QLabel(
            "• Create a ZIP file by selecting multiple Excel files, right-clicking, and choosing 'Compress'\n"
            "• Make sure your Excel files are readable and not password-protected\n"
            "• Avoid using special characters in filenames"
        )
        tips_label.setWordWrap(True)
        tips_layout.addWidget(tips_label)
        tips_group.setLayout(tips_layout)
        layout.addWidget(tips_group)
        
        # File selection
        file_layout = QHBoxLayout()
        self.file_path_label = QLineEdit()
        self.file_path_label.setReadOnly(True)
        self.file_path_label.setPlaceholderText("No file selected")
        browse_button = QPushButton("Browse...")
        browse_button.clicked.connect(self.browse_zip_file)
        file_layout.addWidget(self.file_path_label)
        file_layout.addWidget(browse_button)
        layout.addLayout(file_layout)
        
        # Process button
        process_button = QPushButton("Process ZIP File")
        process_button.clicked.connect(self.process_zip_file)
        layout.addWidget(process_button)
        
        # Log area
        log_group = QGroupBox("Processing Log")
        log_layout = QVBoxLayout()
        self.log_label = QLabel("No processing log yet")
        self.log_label.setAlignment(Qt.AlignTop)
        self.log_label.setWordWrap(True)
        
        log_scroll = QScrollArea()
        log_scroll.setWidget(self.log_label)
        log_scroll.setWidgetResizable(True)
        log_scroll.setMinimumHeight(200)
        
        log_layout.addWidget(log_scroll)
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)
        
        # Add stretch to position elements
        layout.addStretch()
        
    def setup_selection_tab(self, file_data):
        """Setup UI for the data selection tab based on loaded files using a tree view"""
        # Create a new widget for the selection tab
        selection_widget = QWidget()
        
        # Create main layout for the selection tab
        selection_layout = QVBoxLayout(selection_widget)
        
        # Main content area widget (will contain tree and sheet stack)
        content_widget = QWidget()
        content_layout = QHBoxLayout(content_widget)
        content_layout.setContentsMargins(0, 0, 0, 0)
        
        # Create a splitter for the tree view and content area
        splitter = QSplitter(Qt.Horizontal)
        
        # Create tree view for file and sheet navigation
        self.tree_view = QTreeWidget()
        self.tree_view.setHeaderLabel("Files and Sheets")
        self.tree_view.setMinimumWidth(250)
        self.tree_view.setExpandsOnDoubleClick(True)
        self.tree_view.itemClicked.connect(self.on_tree_item_clicked)
        
        # Create stacked widget for content (will show sheet data and column selection)
        self.sheet_stack = QStackedWidget()
        
        # Add tree view and stacked widget to splitter
        splitter.addWidget(self.tree_view)
        splitter.addWidget(self.sheet_stack)
        
        # Set initial splitter sizes
        splitter.setSizes([250, 650])
        
        # Add splitter to the content layout
        content_layout.addWidget(splitter)
        
        # Add the content widget to the main layout
        selection_layout.addWidget(content_widget)
        
        # Create the navigation buttons layout at the bottom
        button_layout = QHBoxLayout()
        
        # Add spacer to push buttons to the right side
        button_layout.addStretch()
        
        # Create the navigation buttons
        back_btn = QPushButton("Back to Upload")
        next_btn = QPushButton("Continue to Output")
        
        back_btn.clicked.connect(lambda: self.tabs.setCurrentIndex(0))
        next_btn.clicked.connect(self.check_selection_and_continue)
        
        # Add buttons to the layout
        button_layout.addWidget(back_btn)
        button_layout.addWidget(next_btn)
        
        # Add the button layout to the bottom of the main layout
        selection_layout.addLayout(button_layout)
        
        # Clear any existing content
        if hasattr(self, 'selection_tab') and isinstance(self.selection_tab, QWidget):
            # If selection_tab is a QTabWidget, just replace it
            self.tabs.removeTab(1)
            self.tabs.insertTab(1, selection_widget, "2. Select Data")
        else:
            self.selection_tab = selection_widget
            self.tabs.removeTab(1)
            self.tabs.insertTab(1, self.selection_tab, "2. Select Data")
        
        # Populate the tree view with files and sheets
        self.populate_tree_view(file_data)
        
    def populate_tree_view(self, file_data):
        """Populate the tree view with files and sheets"""
        self.tree_view.clear()
        self.sheet_stack.setCurrentIndex(0)
        
        # Clear previous dictionaries to avoid confusion with old data
        self.tree_items = {}
        self.sheet_widgets = {}
        
        # Debug: Print the file data structure to understand the hierarchy
        print("\n---- DEBUG: File Data Structure ----")
        file_count = len(file_data)
        print(f"Total files to display: {file_count}")
        
        if file_count == 0:
            print("WARNING: No files in file_data dictionary!")
            return
        
        # Count total sheets for statistics
        total_sheet_count = 0
        
        for file_name, sheets in file_data.items():
            sheet_count = len(sheets)
            total_sheet_count += sheet_count
            print(f"File: {file_name}")
            sheet_names = list(sheets.keys())
            print(f"  Sheets ({sheet_count}): {', '.join(sheet_names)}")
        
        print(f"Total sheets to display: {total_sheet_count}")
        print("-----------------------------------\n")
        
        # We need to control exactly how widgets are added to the stack to prevent offset issues
        # First, create all the widgets for all files and sheets and then add them to the stack in order
        # This avoids any potential offset issues where tree items don't match stack widget indices
        
        # Step 1: Build data structures first without adding to the stack
        file_items = []
        sheet_items = []
        sheet_widgets = []
        
        # Step 2: Create all file and sheet tree items first
        for file_idx, (file_name, sheets) in enumerate(file_data.items()):
            # Create file item and add to the tree
            file_item = QTreeWidgetItem(self.tree_view)
            file_item.setText(0, file_name)
            file_item.setIcon(0, self.style().standardIcon(QStyle.SP_FileIcon))
            file_item.setExpanded(True)
            
            # Store file item in our tracking list
            file_items.append((file_name, file_item))
            
            # Add sheets as child items
            sheet_count = len(sheets)
            print(f"Processing file: {file_name} with {sheet_count} sheets")
            
            for sheet_idx, (sheet_name, df) in enumerate(sheets.items()):
                # Create the sheet tree item
                sheet_item = QTreeWidgetItem(file_item)
                sheet_item.setText(0, sheet_name)
                sheet_item.setIcon(0, self.style().standardIcon(QStyle.SP_FileDialogDetailedView))
                
                # Store references to navigate to this sheet
                sheet_item.file_name = file_name
                sheet_item.sheet_name = sheet_name
                
                # Create the sheet widget for this sheet
                sheet_widget = self.create_sheet_widget(file_name, sheet_name, df)
                
                # Add to our tracking lists
                sheet_key = f"{file_name}_{sheet_name}"
                sheet_items.append((sheet_key, sheet_item))
                sheet_widgets.append((sheet_key, sheet_widget))
                
                print(f"  Created tree item and widget for sheet: {sheet_name} in file: {file_name}")
        
        # Step 3: Now add all widgets to the stack in a controlled order
        print("\n---- Adding widgets to stack in controlled order ----")
        
        # First, add a welcome widget at index 0
        welcome_widget = QWidget()
        welcome_layout = QVBoxLayout(welcome_widget)
        welcome_label = QLabel("Select a sheet from the tree view on the left to view and select data columns.")
        welcome_label.setAlignment(Qt.AlignCenter)
        welcome_layout.addWidget(welcome_label)
        self.sheet_stack.addWidget(welcome_widget)
        
        # Now add all the sheet widgets
        for idx, (sheet_key, widget) in enumerate(sheet_widgets):
            # Add to stack and store the index 
            widget_idx = self.sheet_stack.addWidget(widget)
            self.sheet_widgets[sheet_key] = widget_idx
            print(f"  Added widget {idx+1}/{len(sheet_widgets)}: {sheet_key} at index {widget_idx}")
            
            # Double check the widget is where we expect it
            actual_widget = self.sheet_stack.widget(widget_idx)
            if actual_widget != widget:
                print(f"  ERROR: Widget mismatch at index {widget_idx}!")
        
        # Step 4: Store all tree items for lookup
        for file_name, file_item in file_items:
            self.tree_items[file_name] = file_item
            
        for sheet_key, sheet_item in sheet_items:
            self.tree_items[sheet_key] = sheet_item
            
        # Final verification
        print(f"\nAdded {len(file_items)} files with {len(sheet_items)} sheets to the tree")
        print(f"Added {len(sheet_widgets)} widgets to the stack")
        print("Dump of sheet_widgets dictionary:")
        for key, idx in sorted(self.sheet_widgets.items()):
            print(f"  {key} -> index {idx}")
        print("----------------------------------------------------\n")
        
    def create_sheet_widget(self, file_name, sheet_name, df):
        """Create a widget for displaying sheet data and column selection"""
        sheet_widget = QWidget()
        sheet_layout = QVBoxLayout(sheet_widget)
        
        # Add file and sheet info at the top
        info_label = QLabel(f"File: {file_name} | Sheet: {sheet_name}")
        info_label.setStyleSheet("font-weight: bold; color: #336699;")
        sheet_layout.addWidget(info_label)
        
        # Data preview
        preview_group = QGroupBox("Data Preview")
        preview_layout = QVBoxLayout()
        
        # Create table view
        table_view = QTableView()
        
        # Super-simplified data handling - keep it as basic as possible
        try:
            # Always try to display some data, regardless of its structure
            if df is not None and not df.empty:
                # Take the first few rows of the dataframe as is, without any preprocessing
                # This ensures we display data even if the first rows are blank
                preview_rows = min(10, len(df))
                sample_df = df.head(preview_rows)
                
                # Create the model with the raw data exactly as it was read
                model = PandasTableModel(sample_df)
                
                # Add informative status message
                total_rows = len(df)
                cols = len(df.columns)
                status_label = QLabel(f"Displaying {preview_rows} of {total_rows} rows - {cols} columns")
                status_label.setStyleSheet("color: #666; font-style: italic;")
                preview_layout.addWidget(status_label)
                
                # Add warning if there might be blank rows
                blank_count = df.isna().all(axis=1).sum()
                if blank_count > 0:
                    blank_label = QLabel(f"Note: This sheet contains {blank_count} blank rows which are kept in the data.")
                    blank_label.setStyleSheet("color: #993300; font-style: italic;")
                    preview_layout.addWidget(blank_label)
            else:
                # Only if truly empty, show a message
                model = PandasTableModel(pd.DataFrame({
                    'Note': ['No data found in this sheet - it appears to be empty']
                }))
        except Exception as e:
            # Handle any errors that might occur
            model = PandasTableModel(pd.DataFrame({
                'Error': [f'Could not display sheet data: {str(e)}']
            }))
            
            # Log the error for debugging
            print(f"Error displaying sheet data: {str(e)}")
        
        # Apply the model to the table view
        table_view.setModel(model)
        
        # Set table properties
        table_view.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table_view.setAlternatingRowColors(False)  # Turn off alternating colors per user request
        
        preview_layout.addWidget(table_view)
        preview_group.setLayout(preview_layout)
        sheet_layout.addWidget(preview_group)
        
        # Column selection
        selection_group = QGroupBox("Select Columns to Extract")
        selection_layout = QVBoxLayout()
        
        # Buttons for select all/none
        buttons_layout = QHBoxLayout()
        select_all_btn = QPushButton("Select All")
        deselect_all_btn = QPushButton("Deselect All")
        
        # Store references to these buttons with the file and sheet info
        select_all_btn.file_name = file_name
        select_all_btn.sheet_name = sheet_name
        deselect_all_btn.file_name = file_name
        deselect_all_btn.sheet_name = sheet_name
        
        select_all_btn.clicked.connect(self.select_all_columns)
        deselect_all_btn.clicked.connect(self.deselect_all_columns)
        
        buttons_layout.addWidget(select_all_btn)
        buttons_layout.addWidget(deselect_all_btn)
        selection_layout.addLayout(buttons_layout)
        
        # Column checkboxes
        scroll_widget = QWidget()
        scroll_layout = QGridLayout(scroll_widget)
        
        # Make sure the selected_columns structure is initialized
        if file_name not in self.selected_columns:
            self.selected_columns[file_name] = {}
        if sheet_name not in self.selected_columns[file_name]:
            self.selected_columns[file_name][sheet_name] = []
        
        # Add checkboxes in a grid
        columns = df.columns.tolist()
        cols_per_row = 3
        for i, col_name in enumerate(columns):
            row = i // cols_per_row
            col = i % cols_per_row
            
            checkbox = QCheckBox(str(col_name))
            
            # Store column info in the checkbox
            checkbox.file_name = file_name
            checkbox.sheet_name = sheet_name
            checkbox.column_name = col_name
            
            checkbox.stateChanged.connect(self.column_selection_changed)
            scroll_layout.addWidget(checkbox, row, col)
        
        # Adjust grid layout
        scroll_layout.setColumnStretch(0, 1)
        scroll_layout.setColumnStretch(1, 1)
        scroll_layout.setColumnStretch(2, 1)
        
        # Create scroll area for checkboxes
        scroll_area = QScrollArea()
        scroll_area.setWidget(scroll_widget)
        scroll_area.setWidgetResizable(True)
        selection_layout.addWidget(scroll_area)
        
        selection_group.setLayout(selection_layout)
        sheet_layout.addWidget(selection_group)
        
        return sheet_widget
        
    def on_tree_item_clicked(self, item, column):
        """Handle tree view item click to display the corresponding sheet"""
        # First, let's just ensure proper labels in the UI
        if hasattr(item, 'file_name') and hasattr(item, 'sheet_name'):
            # This item is a sheet, not a file
            file_name = item.file_name
            sheet_name = item.sheet_name
            key = f"{file_name}_{sheet_name}"
            
            # Enhanced debug information to help diagnose issues
            print(f"\n================ DETAILED SHEET SELECTION DEBUG ================")
            print(f"CLICKED: file={file_name}, sheet={sheet_name}, key={key}")
            print(f"TREE ITEM TEXT: {item.text(0)}")
            print(f"TREE ITEM PARENT: {item.parent().text(0) if item.parent() else 'ROOT'}")
            
            # Get information about all sheets in this file (from our data structure)
            if file_name in self.file_data:
                sheets_in_file = list(self.file_data[file_name].keys())
                print(f"SHEETS IN THIS FILE: {sheets_in_file}")
                if sheet_name in sheets_in_file:
                    print(f"✓ Sheet '{sheet_name}' exists in file '{file_name}' data")
                else:
                    print(f"✗ Sheet '{sheet_name}' NOT found in file '{file_name}' data!")
            else:
                print(f"✗ File '{file_name}' NOT found in file_data!")
            
            # Check if we have a widget for this sheet
            print(f"\nLooking for widget with key: {key}")
            matching_keys = [k for k in self.sheet_widgets.keys() if k == key]
            print(f"Exact matching keys: {matching_keys}")
            
            # Print full mapping for debugging
            print("\nFull sheet_widgets mapping:")
            for k, idx in sorted(self.sheet_widgets.items()):
                parts = k.split('_')
                if len(parts) > 1:
                    file = parts[0]
                    if file == file_name:
                        print(f"  ➤ {k} -> index {idx}")
                    else:
                        print(f"    {k} -> index {idx}")
            
            # Now attempt to find and display the widget
            if key in self.sheet_widgets:
                widget_idx = self.sheet_widgets[key]
                print(f"\n✓ SUCCESS: Found widget at index {widget_idx}")
                
                # Get the widget to verify it's the right one
                widget = self.sheet_stack.widget(widget_idx)
                info_label = None
                for child in widget.findChildren(QLabel):
                    if "File:" in child.text():
                        info_label = child
                        break
                
                if info_label:
                    print(f"Widget header: {info_label.text()}")
                
                # Show the sheet widget
                self.sheet_stack.setCurrentIndex(widget_idx)
            else:
                print(f"\n✗ ERROR: Widget with key {key} not found in sheet_widgets dictionary!")
                
                # Emergency recovery procedure
                # First try to find a sheet with the same name
                sheet_matches = [k for k in self.sheet_widgets.keys() if f"_{sheet_name}" in k]
                print(f"Potential sheet name matches: {sheet_matches}")
                
                # Try to find a sheet in the same file
                file_matches = [k for k in self.sheet_widgets.keys() if k.startswith(f"{file_name}_")]
                print(f"Potential file matches: {file_matches}")
                
                best_match = None
                
                # Prioritize file matches first (sheets in the same file)
                if file_matches:
                    best_match = file_matches[0]
                    print(f"Using match from same file: {best_match}")
                # Then try sheets with the same name
                elif sheet_matches:
                    best_match = sheet_matches[0]
                    print(f"Using match with same sheet name: {best_match}")
                
                # If we found any match, use it as an emergency fallback
                if best_match:
                    widget_idx = self.sheet_widgets[best_match]
                    print(f"Emergency display: Using widget at index {widget_idx}")
                    self.sheet_stack.setCurrentIndex(widget_idx)
                    
                    # Show a popup warning to the user
                    QMessageBox.warning(
                        self, 
                        "Sheet Display Issue", 
                        f"There was an issue displaying the exact sheet you selected.\n" +
                        f"Showing a similar sheet instead as an emergency fallback.\n\n" +
                        f"Selected: {file_name} - {sheet_name}\n" +
                        f"Showing: {best_match}\n\n" +
                        f"Please contact support with this information."
                    )
                else:
                    print("No matching sheet found - staying on current view")
            
            print("================ END SHEET SELECTION DEBUG ================\n")
        
    def setup_output_tab(self):
        """Setup UI for the output tab"""
        layout = QVBoxLayout(self.output_tab)
        
        # Instructions
        instruction_label = QLabel(
            "Enter a name for the output Excel file and select where to save it.\n"
            "Click 'Process and Generate' to create the merged Excel file."
        )
        instruction_label.setWordWrap(True)
        layout.addWidget(instruction_label)
        
        # Output file name
        name_layout = QHBoxLayout()
        name_label = QLabel("Output filename:")
        self.output_name_edit = QLineEdit("merged_data")
        name_layout.addWidget(name_label)
        name_layout.addWidget(self.output_name_edit)
        layout.addLayout(name_layout)
        
        # Output location
        path_layout = QHBoxLayout()
        path_label = QLabel("Save location:")
        self.output_path_label = QLineEdit()
        self.output_path_label.setReadOnly(True)
        self.output_path_label.setPlaceholderText("No location selected")
        browse_output_button = QPushButton("Browse...")
        browse_output_button.clicked.connect(self.browse_output_location)
        path_layout.addWidget(path_label)
        path_layout.addWidget(self.output_path_label)
        path_layout.addWidget(browse_output_button)
        layout.addLayout(path_layout)
        
        # Process button
        process_output_button = QPushButton("Process and Generate Excel File")
        process_output_button.clicked.connect(self.generate_output_file)
        layout.addWidget(process_output_button)
        
        # Log area
        output_log_group = QGroupBox("Processing Log")
        output_log_layout = QVBoxLayout()
        self.output_log_label = QLabel("No processing log yet")
        self.output_log_label.setAlignment(Qt.AlignTop)
        self.output_log_label.setWordWrap(True)
        
        output_log_scroll = QScrollArea()
        output_log_scroll.setWidget(self.output_log_label)
        output_log_scroll.setWidgetResizable(True)
        output_log_scroll.setMinimumHeight(200)
        
        output_log_layout.addWidget(output_log_scroll)
        output_log_group.setLayout(output_log_layout)
        layout.addWidget(output_log_group)
        
        # Navigation buttons
        nav_layout = QHBoxLayout()
        back_btn = QPushButton("Back to Selection")
        back_btn.clicked.connect(lambda: self.tabs.setCurrentIndex(1))
        nav_layout.addWidget(back_btn)
        layout.addLayout(nav_layout)
        
        # Add stretch to position elements
        layout.addStretch()
    
    def browse_zip_file(self):
        """Open file dialog to select ZIP file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select ZIP File", "", "ZIP Files (*.zip)"
        )
        
        if file_path:
            self.file_path_label.setText(file_path)
    
    def browse_output_location(self):
        """Open file dialog to select output location"""
        # MacOS-style save dialog
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save Output Excel File", 
            f"{self.output_name_edit.text()}.xls",
            "Excel Files (*.xls)"
        )
        
        if file_path:
            # Update the path label and extract just the directory
            self.output_path_label.setText(file_path)
            self.output_path = file_path
    
    def process_zip_file(self):
        """Process the selected ZIP file"""
        zip_path = self.file_path_label.text()
        
        if not zip_path:
            QMessageBox.warning(
                self, "No File Selected", 
                "Please select a ZIP file containing Excel files."
            )
            return
        
        if not os.path.exists(zip_path):
            QMessageBox.warning(
                self, "File Not Found", 
                f"The selected file does not exist: {zip_path}"
            )
            return
        
        # Create a temporary directory for extraction
        self.temp_dir = tempfile.mkdtemp()
        
        # Clear the log
        self.log_label.setText("")
        
        # Show progress bar
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        
        # Disable the tab during processing
        self.tabs.setTabEnabled(0, False)
        
        # Create and start the worker thread
        self.file_processor = FileProcessorThread(zip_path, self.temp_dir)
        
        # Connect signals
        self.file_processor.progress_signal.connect(self.update_log)
        self.file_processor.finished_signal.connect(self.processing_finished)
        self.file_processor.error_signal.connect(self.processing_error)
        
        # Start processing
        self.file_processor.start()
    
    def update_log(self, message):
        """Update the log with new message"""
        current_text = self.log_label.text()
        if current_text == "No processing log yet":
            current_text = ""
        
        new_text = current_text + message + "\n"
        self.log_label.setText(new_text)
        
        # Update progress indicators
        self.statusBar().showMessage(message)
        self.progress_bar.setValue((self.progress_bar.value() + 5) % 100)  # Simple animation
    
    def processing_finished(self, file_data):
        """Handle successful processing of ZIP file"""
        # Hide progress bar
        self.progress_bar.setVisible(False)
        
        # Store the data
        self.file_data = file_data
        
        # Setup the selection tab with the file data
        self.setup_selection_tab(file_data)
        
        # Enable the selection tab and switch to it
        self.tabs.setTabEnabled(0, True)
        self.tabs.setTabEnabled(1, True)
        self.tabs.setCurrentIndex(1)
        
        # Update status
        self.statusBar().showMessage("Ready to select data")
    
    def processing_error(self, error_message):
        """Handle error during processing"""
        # Hide progress bar
        self.progress_bar.setVisible(False)
        
        # Show error message
        QMessageBox.critical(
            self, "Processing Error", 
            f"An error occurred while processing the ZIP file:\n{error_message}"
        )
        
        # Update status and re-enable tab
        self.statusBar().showMessage("Error processing ZIP file")
        self.tabs.setTabEnabled(0, True)
    
    def column_selection_changed(self, state):
        """Handle column selection checkbox changes"""
        # Get the sender checkbox
        checkbox = self.sender()
        
        file_name = checkbox.file_name
        sheet_name = checkbox.sheet_name
        column_name = checkbox.column_name
        
        # Update the selected columns structure
        if state == Qt.Checked:
            if column_name not in self.selected_columns[file_name][sheet_name]:
                self.selected_columns[file_name][sheet_name].append(column_name)
        else:
            if column_name in self.selected_columns[file_name][sheet_name]:
                self.selected_columns[file_name][sheet_name].remove(column_name)
        
        # Update status with selection count
        total_selected = sum(
            len(cols) for file in self.selected_columns.values() 
            for cols in file.values()
        )
        self.statusBar().showMessage(f"Selected {total_selected} columns")
    
    def select_all_columns(self):
        """Select all columns for a sheet"""
        # Get the sender button
        button = self.sender()
        
        file_name = button.file_name
        sheet_name = button.sheet_name
        
        # Get all columns for this sheet
        all_columns = self.file_data[file_name][sheet_name].columns.tolist()
        
        # Update the selected columns structure
        self.selected_columns[file_name][sheet_name] = all_columns.copy()
        
        # Update checkboxes
        self.update_checkboxes_for_sheet(file_name, sheet_name)
        
        # Update status
        total_selected = sum(
            len(cols) for file in self.selected_columns.values() 
            for cols in file.values()
        )
        self.statusBar().showMessage(f"Selected {total_selected} columns")
    
    def deselect_all_columns(self):
        """Deselect all columns for a sheet"""
        # Get the sender button
        button = self.sender()
        
        file_name = button.file_name
        sheet_name = button.sheet_name
        
        # Clear the selected columns for this sheet
        self.selected_columns[file_name][sheet_name] = []
        
        # Update checkboxes
        self.update_checkboxes_for_sheet(file_name, sheet_name)
        
        # Update status
        total_selected = sum(
            len(cols) for file in self.selected_columns.values() 
            for cols in file.values()
        )
        self.statusBar().showMessage(f"Selected {total_selected} columns")
    
    def update_checkboxes_for_sheet(self, file_name, sheet_name):
        """Update all checkboxes for a specific sheet to match selection state"""
        # Find the sheet widget in our stacked widget
        key = f"{file_name}_{sheet_name}"
        if key in self.sheet_widgets:
            widget_idx = self.sheet_widgets[key]
            sheet_widget = self.sheet_stack.widget(widget_idx)
            
            # Find the QScrollArea in the second QGroupBox (column selection)
            groups = sheet_widget.findChildren(QGroupBox)
            if len(groups) >= 2:
                selection_group = groups[1]  # Second group is selection
                scroll_area = selection_group.findChild(QScrollArea)
                
                if scroll_area and scroll_area.widget():
                    scroll_widget = scroll_area.widget()
                    
                    # Update all checkboxes
                    for checkbox in scroll_widget.findChildren(QCheckBox):
                        if (hasattr(checkbox, 'file_name') and 
                            hasattr(checkbox, 'sheet_name') and 
                            hasattr(checkbox, 'column_name')):
                            if (checkbox.file_name == file_name and 
                                checkbox.sheet_name == sheet_name):
                                # Block signals to prevent recursive calls
                                checkbox.blockSignals(True)
                                checkbox.setChecked(
                                    checkbox.column_name in self.selected_columns[file_name][sheet_name]
                                )
                                checkbox.blockSignals(False)
    
    def check_selection_and_continue(self):
        """Check if any columns are selected before continuing"""
        total_selected = sum(
            len(cols) for file in self.selected_columns.values() 
            for cols in file.values()
        )
        
        if total_selected == 0:
            QMessageBox.warning(
                self, "No Columns Selected", 
                "Please select at least one column to extract."
            )
            return
        
        # Enable the output tab and switch to it
        self.tabs.setTabEnabled(2, True)
        self.tabs.setCurrentIndex(2)
    
    def generate_output_file(self):
        """Generate the output Excel file"""
        # Check if output path is selected
        if not self.output_path:
            # Try to get a path now
            self.browse_output_location()
            if not self.output_path:
                QMessageBox.warning(
                    self, "No Output Location", 
                    "Please select a location to save the output file."
                )
                return
        
        # Clear the output log
        self.output_log_label.setText("")
        
        # Show progress bar
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        
        # Disable the tab during processing
        self.tabs.setTabEnabled(2, False)
        
        # Create and start the worker thread
        self.output_processor = OutputProcessorThread(
            self.file_data, 
            self.selected_columns,
            self.output_path
        )
        
        # Connect signals
        self.output_processor.progress_signal.connect(self.update_output_log)
        self.output_processor.finished_signal.connect(self.output_finished)
        self.output_processor.error_signal.connect(self.output_error)
        
        # Start processing
        self.output_processor.start()
    
    def update_output_log(self, message):
        """Update the output log with new message"""
        current_text = self.output_log_label.text()
        if current_text == "No processing log yet":
            current_text = ""
        
        new_text = current_text + message + "\n"
        self.output_log_label.setText(new_text)
        
        # Update progress indicators
        self.statusBar().showMessage(message)
        self.progress_bar.setValue((self.progress_bar.value() + 5) % 100)  # Simple animation
    
    def output_finished(self, output_path):
        """Handle successful generation of output file"""
        # Hide progress bar
        self.progress_bar.setVisible(False)
        
        # Show success message
        QMessageBox.information(
            self, "Processing Complete", 
            f"The merged Excel file has been saved to:\n{output_path}"
        )
        
        # Update status and re-enable tab
        self.statusBar().showMessage("Processing complete")
        self.tabs.setTabEnabled(2, True)
        
        # Ask if user wants to process another file
        reply = QMessageBox.question(
            self, "Process Another?", 
            "Would you like to process another ZIP file?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.reset_app()
    
    def output_error(self, error_message):
        """Handle error during output generation"""
        # Hide progress bar
        self.progress_bar.setVisible(False)
        
        # Show error message
        QMessageBox.critical(
            self, "Processing Error", 
            f"An error occurred while generating the output file:\n{error_message}"
        )
        
        # Update status and re-enable tab
        self.statusBar().showMessage("Error generating output file")
        self.tabs.setTabEnabled(2, True)
    
    def reset_app(self):
        """Reset the application to initial state"""
        # Clear all data
        self.file_data = {}
        self.selected_columns = {}
        self.output_path = None
        
        # Clear UI elements
        self.file_path_label.setText("")
        self.log_label.setText("No processing log yet")
        self.output_path_label.setText("")
        self.output_log_label.setText("No processing log yet")
        self.output_name_edit.setText("merged_data")
        
        # Reset tab states
        self.tabs.setTabEnabled(0, True)
        self.tabs.setTabEnabled(1, False)
        self.tabs.setTabEnabled(2, False)
        self.tabs.setCurrentIndex(0)
        
        # Clean up temporary directory if it exists
        if self.temp_dir and os.path.exists(self.temp_dir):
            try:
                import shutil
                shutil.rmtree(self.temp_dir)
                self.temp_dir = None
            except Exception as e:
                print(f"Error cleaning temporary directory: {str(e)}")
        
        # Update status
        self.statusBar().showMessage("Application reset and ready")
    
    def closeEvent(self, event):
        """Clean up on application close"""
        # Clean up temporary directory
        if self.temp_dir and os.path.exists(self.temp_dir):
            try:
                import shutil
                shutil.rmtree(self.temp_dir)
            except Exception as e:
                print(f"Error cleaning temporary directory: {e}")
        
        event.accept()

def main():
    # Set macOS-specific application attributes
    if sys.platform == 'darwin':
        # Set application ID
        QApplication.setApplicationName("Excel Data Extractor")
        QApplication.setOrganizationName("MacOS Excel Tools")
        QApplication.setOrganizationDomain("macostools.example.com")
        
        # High DPI scaling for Retina displays
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    
    app = QApplication(sys.argv)
    
    # Apply styling based on platform
    if sys.platform == 'darwin':
        # Use macOS native style for best integration
        app.setStyle("macintosh")
    else:
        # For non-macOS platforms, use Fusion style which looks modern
        app.setStyle("Fusion")
    
    window = ExcelExtractorApp()
    window.show()
    
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()