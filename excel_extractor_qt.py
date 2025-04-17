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
    QLineEdit, QTableView, QHeaderView, QSplitter, QFrame, QStyle
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QAbstractTableModel, QModelIndex, QSize
from PyQt5.QtGui import QFont, QIcon, QPalette, QColor

# Model for displaying Excel data in a table
class PandasTableModel(QAbstractTableModel):
    def __init__(self, data):
        super().__init__()
        self._data = data

    def rowCount(self, parent=None):
        return len(self._data)

    def columnCount(self, parent=None):
        return len(self._data.columns)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        if role == Qt.DisplayRole:
            value = self._data.iloc[index.row(), index.column()]
            return str(value)
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                return str(self._data.columns[section])
            else:
                return str(section + 1)
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
        
        try:
            self.progress_signal.emit("Opening ZIP file...")
            
            with zipfile.ZipFile(self.zip_path, 'r') as zip_ref:
                # List all files in the ZIP
                file_list = zip_ref.namelist()
                
                self.progress_signal.emit(f"Found {len(file_list)} files in ZIP archive")
                
                # Extract only Excel files
                for file_name in file_list:
                    lower_name = file_name.lower()
                    if lower_name.endswith('.xlsx') or lower_name.endswith('.xls'):
                        # Handle folder paths in ZIP
                        if file_name.endswith('/') or os.path.basename(file_name) == '':
                            continue
                            
                        # Extract the file
                        try:
                            self.progress_signal.emit(f"Extracting: {file_name}")
                            zip_ref.extract(file_name, self.extract_dir)
                            full_path = os.path.join(self.extract_dir, file_name)
                            excel_files.append(full_path)
                        except Exception as extract_error:
                            self.progress_signal.emit(f"Could not extract {file_name}: {str(extract_error)}")
                    
                # Also look for Excel files in extracted folders
                for root, dirs, files in os.walk(self.extract_dir):
                    for file in files:
                        if file.lower().endswith(('.xlsx', '.xls')) and os.path.join(root, file) not in excel_files:
                            excel_files.append(os.path.join(root, file))
                            self.progress_signal.emit(f"Found additional Excel file: {file}")
        
        except Exception as e:
            self.error_signal.emit(f"Error extracting ZIP file: {str(e)}")
            return []
        
        self.progress_signal.emit(f"Extracted {len(excel_files)} Excel files")
        return excel_files
    
    def read_excel_files(self, file_paths):
        """Read data from multiple Excel files"""
        file_data = {}
        
        if not file_paths:
            self.progress_signal.emit("No Excel files to process")
            return file_data
        
        self.progress_signal.emit(f"Reading {len(file_paths)} Excel files...")
        
        for file_path in file_paths:
            try:
                # Get just the filename without path
                file_name = os.path.basename(file_path)
                self.progress_signal.emit(f"Reading: {file_name}")
                
                # Read all sheets from the Excel file
                try:
                    excel_file = pd.ExcelFile(file_path)
                    sheet_names = excel_file.sheet_names
                    self.progress_signal.emit(f"Found {len(sheet_names)} sheets in {file_name}")
                except Exception as excel_error:
                    self.progress_signal.emit(f"Error opening Excel file '{file_name}': {str(excel_error)}")
                    
                    # Try alternate approach for older Excel formats
                    try:
                        # For xls files
                        if file_path.lower().endswith('.xls'):
                            df = pd.read_excel(file_path, engine='xlrd')
                            file_data[file_name] = {"Sheet1": df}
                            self.progress_signal.emit(f"Successfully read {file_name} using xlrd engine")
                            continue
                    except Exception as alt_error:
                        self.progress_signal.emit(f"Alternative read approach failed: {str(alt_error)}")
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
                            self.progress_signal.emit(f"Sheet '{sheet_name}' has {len(df)} rows and {len(df.columns)} columns")
                        else:
                            self.progress_signal.emit(f"Sheet '{sheet_name}' is empty, skipping")
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
        """Setup UI for the data selection tab based on loaded files"""
        # Clear any existing tabs
        while self.selection_tab.count() > 0:
            self.selection_tab.removeTab(0)
        
        # Create a tab for each Excel file
        for file_name, sheets in file_data.items():
            file_tab = QWidget()
            file_layout = QVBoxLayout(file_tab)
            
            # Create a nested tab widget for each sheet in the file
            sheet_tabs = QTabWidget()
            file_layout.addWidget(sheet_tabs)
            
            # Add a tab for each sheet
            for sheet_name, df in sheets.items():
                sheet_tab = QWidget()
                sheet_layout = QVBoxLayout(sheet_tab)
                
                # Data preview
                preview_group = QGroupBox("Data Preview")
                preview_layout = QVBoxLayout()
                
                # Create table view
                table_view = QTableView()
                model = PandasTableModel(df.head(5))  # Show first 5 rows
                table_view.setModel(model)
                
                # Set table properties
                table_view.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
                table_view.setAlternatingRowColors(True)
                
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
                
                sheet_tabs.addTab(sheet_tab, sheet_name)
            
            self.selection_tab.addTab(file_tab, file_name)
        
        # Add navigation buttons at the bottom of each file tab
        for i in range(self.selection_tab.count()):
            file_tab = self.selection_tab.widget(i)
            file_layout = file_tab.layout()
            
            nav_layout = QHBoxLayout()
            back_btn = QPushButton("Back to Upload")
            next_btn = QPushButton("Continue to Output")
            
            back_btn.clicked.connect(lambda: self.tabs.setCurrentIndex(0))
            next_btn.clicked.connect(self.check_selection_and_continue)
            
            nav_layout.addWidget(back_btn)
            nav_layout.addWidget(next_btn)
            file_layout.addLayout(nav_layout)
        
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
        # Find the file tab
        for i in range(self.selection_tab.count()):
            if self.selection_tab.tabText(i) == file_name:
                file_tab = self.selection_tab.widget(i)
                
                # Find the sheet tab within the file tab
                sheet_tabs = file_tab.findChild(QTabWidget)
                for j in range(sheet_tabs.count()):
                    if sheet_tabs.tabText(j) == sheet_name:
                        sheet_tab = sheet_tabs.widget(j)
                        
                        # Find the QScrollArea in the second QGroupBox
                        groups = sheet_tab.findChildren(QGroupBox)
                        selection_group = groups[1]  # Second group is selection
                        scroll_area = selection_group.findChild(QScrollArea)
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
                        
                        break
                break
    
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