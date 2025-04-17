"""
Excel Data Extractor - PyQt5 macOS Application
This application extracts and merges selected data from multiple Excel files in a ZIP archive.
Optimized specifically for macOS with native look and feel.
"""

import os
import sys
import tempfile
import zipfile
import shutil
import logging
from pathlib import Path

import pandas as pd
import numpy as np
import xlwt
import openpyxl

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QMessageBox, QProgressBar,
    QTabWidget, QCheckBox, QGroupBox, QScrollArea, QGridLayout,
    QLineEdit, QTableView, QHeaderView, QSplitter, QFrame, QStyle,
    QTreeWidget, QTreeWidgetItem, QStackedWidget, QComboBox, QDialog,
    QMenuBar, QMenu, QAction
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QAbstractTableModel, QModelIndex, QSize
from PyQt5.QtGui import QFont, QIcon, QPalette, QColor

# Import profile management
try:
    from profile_manager import ProfileManager, ExtractionProfile
    from profile_dialog import ProfileDialog
    PROFILE_SUPPORT = True
except ImportError:
    print("Profile management modules not found, profile support will be disabled")
    PROFILE_SUPPORT = False

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
        
        # Dark mode compatibility: Use system palette for proper colors
        # Don't set any explicit background colors to respect OS theme
        
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
                            from file_processor import detect_descriptive_column_names
                            self.progress_signal.emit(f"Performing advanced header detection for {sheet_name}")
                            
                            # Store the data in our main dictionary
                            file_data[file_name][sheet_name] = temp_excel_file[file_name][sheet_name]
                            
                        except ImportError:
                            # Fallback to direct sheet reading if the shared module is not available
                            self.progress_signal.emit(f"WARNING: Using legacy header detection for {sheet_name}")
                            file_data[file_name][sheet_name] = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
                            
                    except Exception as sheet_error:
                        self.progress_signal.emit(f"Error reading sheet '{sheet_name}' in file '{file_name}': {str(sheet_error)}")
                        # Skip this sheet but continue with others
                        continue
                        
            except Exception as file_error:
                self.progress_signal.emit(f"Error processing file {file_path}: {str(file_error)}")
                # Continue with next file
                continue
        
        # Verify all files were processed
        for original_filename in original_filenames:
            if original_filename not in processed_files:
                self.progress_signal.emit(f"WARNING: File {original_filename} was not processed!")
        
        # Final report
        total_processed = len(file_data)
        total_sheets = sum(len(sheets) for sheets in file_data.values())
        
        self.progress_signal.emit(f"Successfully read {total_processed} files with {total_sheets} sheets")
        
        return file_data

# Worker thread for generating output
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
            # Generate the merged Excel file
            success = self.process_and_merge_data()
            
            if success:
                self.finished_signal.emit(self.output_path)
            else:
                self.error_signal.emit("Error generating merged Excel file")
                
        except Exception as e:
            self.error_signal.emit(f"Error generating output: {str(e)}")
            
    def process_and_merge_data(self):
        """Process and merge selected data from multiple Excel files"""
        try:
            self.progress_signal.emit("Starting data processing and merging...")
            
            # Use the shared processing logic
            try:
                from file_processor import process_and_merge_data
                success = process_and_merge_data(
                    self.file_data, 
                    self.selected_columns, 
                    self.output_path,
                    log_callback=lambda msg: self.progress_signal.emit(msg)
                )
                return success
            except ImportError:
                self.progress_signal.emit("ERROR: Could not import shared processor module")
                return False
                
        except Exception as e:
            self.error_signal.emit(f"Error in processing and merging: {str(e)}")
            return False

# Main application window
class ExcelExtractorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # Initialize data structures
        self.file_data = {}
        self.selected_columns = {}
        self.output_path = None
        self.tree_items = {}
        self.sheet_widgets = {}
        
        # Create temporary directory for extracted files
        self.temp_dir = tempfile.mkdtemp()
        
        # Initialize profile manager if supported
        if PROFILE_SUPPORT:
            try:
                self.profile_manager = ProfileManager()
                self.profile_manager.load_all_profiles()
                self.profile_manager.load_settings()
            except Exception as e:
                print(f"Error initializing profile manager: {str(e)}")
                self.profile_manager = None
        else:
            self.profile_manager = None
                
        # Setup UI
        self.init_ui()
        
    def init_ui(self):
        # Set window properties with macOS optimizations
        self.setWindowTitle("Excel Data Extractor")
        self.setGeometry(100, 100, 900, 600)
        
        # Set application icon using system icon (document icon on macOS)
        self.setWindowIcon(self.style().standardIcon(QStyle.SP_FileDialogDetailedView))
        
        # Create menu bar with profile management
        if PROFILE_SUPPORT and self.profile_manager:
            self.create_menu_bar()
        
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
        
    def create_menu_bar(self):
        """Create the application menu bar"""
        # Create the menu bar
        menu_bar = self.menuBar()
        
        # File menu
        file_menu = menu_bar.addMenu("File")
        
        # Open ZIP action
        open_action = QAction("Open ZIP File...", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self.browse_zip_file)
        file_menu.addAction(open_action)
        
        # Save action
        save_action = QAction("Save Output...", self)
        save_action.setShortcut("Ctrl+S")
        save_action.triggered.connect(self.browse_output_location)
        file_menu.addAction(save_action)
        
        file_menu.addSeparator()
        
        # Exit action
        exit_action = QAction("Exit", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # Profiles menu
        profiles_menu = menu_bar.addMenu("Profiles")
        
        # Manage profiles action
        manage_profiles_action = QAction("Manage Profiles...", self)
        manage_profiles_action.triggered.connect(self.open_profile_manager)
        profiles_menu.addAction(manage_profiles_action)
        
        # Add profile actions for each existing profile
        if self.profile_manager and self.profile_manager.get_all_profiles():
            profiles_menu.addSeparator()
            
            for name, profile in sorted(self.profile_manager.get_all_profiles().items()):
                profile_action = QAction(name, self)
                profile_action.triggered.connect(lambda checked, p=profile: self.apply_profile(p))
                
                # Mark default profile
                if name == self.profile_manager.default_profile_name:
                    profile_action.setText(f"{name} (Default)")
                    
                profiles_menu.addAction(profile_action)
        
        # Help menu
        help_menu = menu_bar.addMenu("Help")
        
        # About action
        about_action = QAction("About Excel Data Extractor", self)
        about_action.triggered.connect(self.show_about_dialog)
        help_menu.addAction(about_action)
    
    def show_about_dialog(self):
        """Show the about dialog"""
        QMessageBox.about(
            self,
            "About Excel Data Extractor",
            "<h3>Excel Data Extractor</h3>"
            "<p>A native macOS application for extracting and merging data from Excel files.</p>"
            "<p>Version 1.0</p>"
            "<p>&copy; 2025 macOS Excel Tools</p>"
        )
        
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
        
        # Add profile management if supported
        if PROFILE_SUPPORT and self.profile_manager:
            profile_group = QGroupBox("Extraction Profiles")
            profile_layout = QVBoxLayout()
            
            profile_info = QLabel(
                "Profiles allow you to save column selections for repeated extraction tasks.\n"
                "You can create profiles based on your current selections or load existing profiles."
            )
            profile_info.setWordWrap(True)
            profile_layout.addWidget(profile_info)
            
            # Profile buttons
            profile_btn_layout = QHBoxLayout()
            
            self.manage_profiles_btn = QPushButton("Manage Profiles")
            self.manage_profiles_btn.clicked.connect(self.open_profile_manager)
            profile_btn_layout.addWidget(self.manage_profiles_btn)
            
            # Add default profile dropdown if we have profiles
            if self.profile_manager.get_all_profiles():
                profile_btn_layout.addWidget(QLabel("Default Profile:"))
                self.profile_combo = QComboBox()
                self.update_profile_combo()
                profile_btn_layout.addWidget(self.profile_combo)
            
            profile_layout.addLayout(profile_btn_layout)
            profile_group.setLayout(profile_layout)
            layout.addWidget(profile_group)
        
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
        
        # Get descriptive column names from the first non-empty string in each column
        # (This feature can later be made configurable in settings)
        try:
            from file_processor import detect_descriptive_column_names
            descriptive_names = detect_descriptive_column_names(df, lambda msg: print(f"Column names: {msg}"))
            # Store these descriptive names for later use
            sheet_key = f"{file_name}_{sheet_name}"
            if not hasattr(self, 'descriptive_column_names'):
                self.descriptive_column_names = {}
            self.descriptive_column_names[sheet_key] = descriptive_names
            print(f"Found {len(descriptive_names)} descriptive column names for {sheet_key}")
        except Exception as e:
            print(f"Error detecting descriptive column names: {str(e)}")
            descriptive_names = {col: col for col in df.columns}  # Default to original names
        
        # Data preview
        preview_group = QGroupBox("Data Preview")
        preview_layout = QVBoxLayout()
        
        # Create the table view
        table_view = QTableView()
        
        # Create a model for the table
        model = PandasTableModel(df)
        table_view.setModel(model)
        
        # Configure the table view
        table_view.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        table_view.horizontalHeader().setStretchLastSection(True)
        table_view.verticalHeader().setDefaultSectionSize(24)
        
        # Set the table to take up available space
        table_view.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        
        # Add the table to the preview layout
        preview_layout.addWidget(table_view)
        
        # Set the layout on the preview group
        preview_group.setLayout(preview_layout)
        
        # Add the preview group to the sheet layout
        sheet_layout.addWidget(preview_group)
        
        # Column selection
        selection_group = QGroupBox("Column Selection")
        selection_layout = QVBoxLayout()
        
        # Add select all / deselect all buttons
        button_layout = QHBoxLayout()
        
        select_all_btn = QPushButton("Select All")
        select_all_btn.clicked.connect(self.select_all_columns)
        button_layout.addWidget(select_all_btn)
        
        deselect_all_btn = QPushButton("Deselect All")
        deselect_all_btn.clicked.connect(self.deselect_all_columns)
        button_layout.addWidget(deselect_all_btn)
        
        selection_layout.addLayout(button_layout)
        
        # Create a scroll area for column checkboxes
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        
        # Create a widget to hold the checkboxes
        checkbox_widget = QWidget()
        checkbox_layout = QGridLayout(checkbox_widget)
        
        # Track the number of columns processed
        col_count = len(df.columns)
        print(f"  Creating {col_count} column checkboxes for {file_name}/{sheet_name}")
        
        # Calculate rows and columns for grid layout
        cols_per_row = 3
        rows_needed = (col_count + cols_per_row - 1) // cols_per_row
        
        # Add a checkbox for each column
        for i, col in enumerate(df.columns):
            # Get descriptive name if available
            if col in descriptive_names:
                display_name = descriptive_names[col]
            else:
                display_name = f"Column {col}"
            
            # Create the checkbox
            checkbox = QCheckBox(display_name)
            checkbox.file_name = file_name
            checkbox.sheet_name = sheet_name
            checkbox.column_name = col
            
            # Connect the checkbox to our selection handler
            checkbox.stateChanged.connect(self.column_selection_changed)
            
            # Add to the grid layout
            row = i // cols_per_row
            col_pos = i % cols_per_row
            checkbox_layout.addWidget(checkbox, row, col_pos)
        
        # Set the layout on the checkbox widget
        checkbox_widget.setLayout(checkbox_layout)
        
        # Add the checkbox widget to the scroll area
        scroll_area.setWidget(checkbox_widget)
        
        # Add the scroll area to the selection layout
        selection_layout.addWidget(scroll_area)
        
        # Set the layout on the selection group
        selection_group.setLayout(selection_layout)
        
        # Add the selection group to the sheet layout
        sheet_layout.addWidget(selection_group)
        
        return sheet_widget
        
    def on_tree_item_clicked(self, item, column):
        """Handle tree view item click to display the corresponding sheet"""
        # Check if this is a sheet item or a file item
        if hasattr(item, 'file_name') and hasattr(item, 'sheet_name'):
            # This is a sheet item, show the corresponding sheet
            sheet_key = f"{item.file_name}_{item.sheet_name}"
            if sheet_key in self.sheet_widgets:
                self.sheet_stack.setCurrentIndex(self.sheet_widgets[sheet_key])
                self.update_checkboxes_for_sheet(item.file_name, item.sheet_name)
        else:
            # This is a file item, show its first sheet or expand/collapse
            if item.childCount() > 0:
                # Toggle expanded state
                item.setExpanded(not item.isExpanded())
                
                # If expanded, show the first sheet
                if item.isExpanded() and item.childCount() > 0:
                    first_sheet_item = item.child(0)
                    if hasattr(first_sheet_item, 'file_name') and hasattr(first_sheet_item, 'sheet_name'):
                        sheet_key = f"{first_sheet_item.file_name}_{first_sheet_item.sheet_name}"
                        if sheet_key in self.sheet_widgets:
                            self.sheet_stack.setCurrentIndex(self.sheet_widgets[sheet_key])
                            self.update_checkboxes_for_sheet(first_sheet_item.file_name, first_sheet_item.sheet_name)
    
    def setup_output_tab(self):
        """Setup UI for the output tab"""
        layout = QVBoxLayout(self.output_tab)
        
        # Instructions
        instruction_label = QLabel(
            "Select an output location and generate the merged Excel file.\n"
            "The file will contain all selected columns from all selected sheets."
        )
        instruction_label.setWordWrap(True)
        layout.addWidget(instruction_label)
        
        # Output location selection
        output_group = QGroupBox("Output Location")
        output_layout = QVBoxLayout()
        
        output_file_layout = QHBoxLayout()
        self.output_path_label = QLineEdit()
        self.output_path_label.setReadOnly(True)
        self.output_path_label.setPlaceholderText("No output location selected")
        
        browse_output_button = QPushButton("Browse...")
        browse_output_button.clicked.connect(self.browse_output_location)
        
        output_file_layout.addWidget(self.output_path_label)
        output_file_layout.addWidget(browse_output_button)
        
        output_layout.addLayout(output_file_layout)
        output_group.setLayout(output_layout)
        layout.addWidget(output_group)
        
        # Summary
        summary_group = QGroupBox("Selection Summary")
        summary_layout = QVBoxLayout()
        
        self.summary_label = QLabel("No data selected yet")
        summary_layout.addWidget(self.summary_label)
        
        summary_group.setLayout(summary_layout)
        layout.addWidget(summary_group)
        
        # Process button
        process_button = QPushButton("Process and Generate Excel File")
        process_button.clicked.connect(self.generate_output_file)
        layout.addWidget(process_button)
        
        # Output log
        output_log_group = QGroupBox("Output Log")
        output_log_layout = QVBoxLayout()
        
        self.output_log_label = QLabel("No output log yet")
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
        button_layout = QHBoxLayout()
        
        back_btn = QPushButton("Back to Selection")
        back_btn.clicked.connect(lambda: self.tabs.setCurrentIndex(1))
        
        # Add spacer to push back button to the left side
        button_layout.addWidget(back_btn)
        button_layout.addStretch()
        
        reset_btn = QPushButton("Start New")
        reset_btn.clicked.connect(self.reset_app)
        button_layout.addWidget(reset_btn)
        
        layout.addLayout(button_layout)
        
        # Add stretch to position elements
        layout.addStretch()
        
    def browse_zip_file(self):
        """Open file dialog to select ZIP file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select ZIP File", "", "ZIP Files (*.zip);;All Files (*)"
        )
        
        if file_path:
            self.file_path_label.setText(file_path)
            
    def browse_output_location(self):
        """Open file dialog to select output location"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save Merged Excel File", "", "Excel Files (*.xlsx);;All Files (*)"
        )
        
        if file_path:
            # Add .xlsx extension if not present
            if not file_path.lower().endswith(('.xlsx', '.xls')):
                file_path += '.xlsx'
                
            self.output_path = file_path
            self.output_path_label.setText(file_path)
            
    def process_zip_file(self):
        """Process the selected ZIP file"""
        # Get the ZIP file path
        zip_path = self.file_path_label.text()
        
        if not zip_path:
            QMessageBox.warning(self, "No File Selected", "Please select a ZIP file first.")
            return
            
        if not os.path.exists(zip_path):
            QMessageBox.critical(self, "File Not Found", f"The selected ZIP file does not exist: {zip_path}")
            return
            
        # Clear previous data
        self.file_data = {}
        self.selected_columns = {}
        
        # Clear the log and show processing message
        self.log_label.setText("Processing ZIP file...")
        self.update_log("Starting ZIP file processing...")
        
        # Show progress bar
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)  # Indeterminate progress
        
        # Create a new temporary directory for this run
        if hasattr(self, 'temp_dir') and os.path.exists(self.temp_dir):
            try:
                shutil.rmtree(self.temp_dir)
            except Exception as e:
                self.update_log(f"Warning: Could not clean previous temporary directory: {str(e)}")
                
        self.temp_dir = tempfile.mkdtemp()
        self.update_log(f"Created temporary directory: {self.temp_dir}")
        
        # Process the ZIP file in a separate thread
        self.processor_thread = FileProcessorThread(zip_path, self.temp_dir)
        self.processor_thread.progress_signal.connect(self.update_log)
        self.processor_thread.finished_signal.connect(self.processing_finished)
        self.processor_thread.error_signal.connect(self.processing_error)
        self.processor_thread.start()
        
    def update_log(self, message):
        """Update the log with new message"""
        current_text = self.log_label.text()
        new_text = current_text + "\n" + message if current_text != "Processing ZIP file..." else message
        self.log_label.setText(new_text)
        
        # Ensure the message is visible in the log widget
        # Since we're using QScrollArea, we need to ensure it scrolls to bottom
        
        # Show progress to user
        self.statusBar().showMessage(message)
        QApplication.processEvents()
        
    def processing_finished(self, file_data):
        """Handle successful processing of ZIP file"""
        self.progress_bar.setVisible(False)
        
        if not file_data:
            QMessageBox.warning(
                self, 
                "No Data Found", 
                "No Excel data could be extracted from the ZIP file. Please check the file and try again."
            )
            return
            
        # Store the processed data
        self.file_data = file_data
        
        # Update the log
        self.update_log(f"Finished processing. Found {len(file_data)} files.")
        
        # Setup the selection tab with the processed data
        self.setup_selection_tab(file_data)
        
        # Enable the selection tab and switch to it
        self.tabs.setTabEnabled(1, True)
        self.tabs.setCurrentIndex(1)
        
        # Show success message
        self.statusBar().showMessage("ZIP file processed successfully. Please select columns to extract.")
        
        # If a profile was selected or is set as default, apply it
        if PROFILE_SUPPORT and self.profile_manager:
            default_profile = self.profile_manager.get_default_profile()
            if default_profile:
                self.apply_profile(default_profile)
                self.update_log(f"Applied default profile: {default_profile.name}")
        
    def processing_error(self, error_message):
        """Handle error during processing"""
        self.progress_bar.setVisible(False)
        self.update_log(f"ERROR: {error_message}")
        QMessageBox.critical(self, "Processing Error", f"Error processing ZIP file: {error_message}")
        
    def column_selection_changed(self, state):
        """Handle column selection checkbox changes"""
        checkbox = self.sender()
        
        # Get the file, sheet and column for this checkbox
        file_name = checkbox.file_name
        sheet_name = checkbox.sheet_name
        column_name = checkbox.column_name
        
        # Ensure the file entry exists in the selected columns dictionary
        if file_name not in self.selected_columns:
            self.selected_columns[file_name] = {}
            
        # Ensure the sheet entry exists in the file entry
        if sheet_name not in self.selected_columns[file_name]:
            self.selected_columns[file_name][sheet_name] = []
            
        # Update the selected columns list based on the checkbox state
        if state == Qt.Checked:
            # Add the column to the selected columns list if not already there
            if column_name not in self.selected_columns[file_name][sheet_name]:
                self.selected_columns[file_name][sheet_name].append(column_name)
        else:
            # Remove the column from the selected columns list if it's there
            if column_name in self.selected_columns[file_name][sheet_name]:
                self.selected_columns[file_name][sheet_name].remove(column_name)
                
        # Remove empty entries from the dictionary to keep it clean
        if not self.selected_columns[file_name][sheet_name]:
            del self.selected_columns[file_name][sheet_name]
            
        if not self.selected_columns[file_name]:
            del self.selected_columns[file_name]
            
        # Print current selection for debugging
        print(f"Column selection changed: {file_name}/{sheet_name}/{column_name} -> {state}")
        self.print_current_selection()
        
    def print_current_selection(self):
        """Print the current selection for debugging"""
        print("\n---- Current Selection ----")
        for file_name, sheets in self.selected_columns.items():
            print(f"File: {file_name}")
            for sheet_name, columns in sheets.items():
                print(f"  Sheet: {sheet_name}")
                print(f"    Columns: {columns}")
        print("--------------------------\n")
        
    def select_all_columns(self):
        """Select all columns for a sheet"""
        # Get the currently displayed sheet
        current_idx = self.sheet_stack.currentIndex()
        
        # Skip the welcome widget at index 0
        if current_idx == 0:
            return
            
        # Find the sheet key for this index
        sheet_key = None
        for key, idx in self.sheet_widgets.items():
            if idx == current_idx:
                sheet_key = key
                break
                
        if not sheet_key:
            print("Could not find sheet key for current index")
            return
            
        # Extract file_name and sheet_name from sheet_key
        parts = sheet_key.split('_', 1)
        if len(parts) != 2:
            print(f"Invalid sheet key format: {sheet_key}")
            return
            
        file_name, sheet_name = parts
        
        # Get the sheet widget
        sheet_widget = self.sheet_stack.widget(current_idx)
        
        # Find and check all checkboxes for this sheet
        for checkbox in self.find_checkboxes(sheet_widget):
            if (
                hasattr(checkbox, 'file_name') and 
                hasattr(checkbox, 'sheet_name') and 
                checkbox.file_name == file_name and 
                checkbox.sheet_name == sheet_name
            ):
                checkbox.setChecked(True)
                
    def deselect_all_columns(self):
        """Deselect all columns for a sheet"""
        # Get the currently displayed sheet
        current_idx = self.sheet_stack.currentIndex()
        
        # Skip the welcome widget at index 0
        if current_idx == 0:
            return
            
        # Find the sheet key for this index
        sheet_key = None
        for key, idx in self.sheet_widgets.items():
            if idx == current_idx:
                sheet_key = key
                break
                
        if not sheet_key:
            print("Could not find sheet key for current index")
            return
            
        # Extract file_name and sheet_name from sheet_key
        parts = sheet_key.split('_', 1)
        if len(parts) != 2:
            print(f"Invalid sheet key format: {sheet_key}")
            return
            
        file_name, sheet_name = parts
        
        # Get the sheet widget
        sheet_widget = self.sheet_stack.widget(current_idx)
        
        # Find and uncheck all checkboxes for this sheet
        for checkbox in self.find_checkboxes(sheet_widget):
            if (
                hasattr(checkbox, 'file_name') and 
                hasattr(checkbox, 'sheet_name') and 
                checkbox.file_name == file_name and 
                checkbox.sheet_name == sheet_name
            ):
                checkbox.setChecked(False)
                
    def find_checkboxes(self, parent):
        """Recursively find all checkboxes in a parent widget"""
        checkboxes = []
        
        # Check if the parent widget is a checkbox
        if isinstance(parent, QCheckBox):
            checkboxes.append(parent)
            
        # Get all child widgets of this parent
        for child in parent.findChildren(QWidget):
            # If the child is a checkbox, add it to the list
            if isinstance(child, QCheckBox):
                checkboxes.append(child)
                
        return checkboxes
        
    def update_checkboxes_for_sheet(self, file_name, sheet_name):
        """Update all checkboxes for a specific sheet to match selection state"""
        # Get the sheet key
        sheet_key = f"{file_name}_{sheet_name}"
        
        # Make sure the sheet is in the widgets dictionary
        if sheet_key not in self.sheet_widgets:
            print(f"Sheet key not found in widgets dictionary: {sheet_key}")
            return
            
        # Get the sheet widget
        sheet_idx = self.sheet_widgets[sheet_key]
        sheet_widget = self.sheet_stack.widget(sheet_idx)
        
        # Check if we have selections for this file/sheet
        has_selections = (
            file_name in self.selected_columns and 
            sheet_name in self.selected_columns[file_name]
        )
        
        # Find all checkboxes for this sheet
        for checkbox in self.find_checkboxes(sheet_widget):
            if (
                hasattr(checkbox, 'file_name') and 
                hasattr(checkbox, 'sheet_name') and 
                hasattr(checkbox, 'column_name') and
                checkbox.file_name == file_name and 
                checkbox.sheet_name == sheet_name
            ):
                # Get the column for this checkbox
                column = checkbox.column_name
                
                # Check if this column is in the selected columns
                if has_selections and column in self.selected_columns[file_name][sheet_name]:
                    checkbox.setChecked(True)
                else:
                    checkbox.setChecked(False)
                    
    def check_selection_and_continue(self):
        """Check if any columns are selected before continuing"""
        if not self.selected_columns:
            QMessageBox.warning(
                self, 
                "No Columns Selected", 
                "Please select at least one column to extract before continuing."
            )
            return
            
        # Generate a summary of the selection
        total_columns = 0
        for file_name, sheets in self.selected_columns.items():
            for sheet_name, columns in sheets.items():
                total_columns += len(columns)
                
        # Update the summary label
        self.summary_label.setText(
            f"Selected {total_columns} columns from {len(self.selected_columns)} files.\n\n"
            "The output file will contain all selected columns merged into a single Excel workbook."
        )
        
        # Enable and switch to the output tab
        self.tabs.setTabEnabled(2, True)
        self.tabs.setCurrentIndex(2)
        
    def generate_output_file(self):
        """Generate the output Excel file"""
        # Check if an output path has been selected
        if not self.output_path:
            QMessageBox.warning(
                self, 
                "No Output Location", 
                "Please select an output location for the merged Excel file."
            )
            self.browse_output_location()
            
            if not self.output_path:
                return
                
        # Check if any columns are selected
        if not self.selected_columns:
            QMessageBox.warning(
                self, 
                "No Columns Selected", 
                "No columns are selected for extraction. Please go back and select columns."
            )
            return
            
        # Clear the output log and show processing message
        self.output_log_label.setText("Generating output file...")
        self.update_output_log("Starting output file generation...")
        
        # Show progress bar
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)  # Indeterminate progress
        
        # Generate the output file in a separate thread
        self.output_thread = OutputProcessorThread(
            self.file_data, self.selected_columns, self.output_path
        )
        self.output_thread.progress_signal.connect(self.update_output_log)
        self.output_thread.finished_signal.connect(self.output_finished)
        self.output_thread.error_signal.connect(self.output_error)
        self.output_thread.start()
        
    def update_output_log(self, message):
        """Update the output log with new message"""
        current_text = self.output_log_label.text()
        new_text = current_text + "\n" + message if current_text != "Generating output file..." else message
        self.output_log_label.setText(new_text)
        
        # Show progress to user
        self.statusBar().showMessage(message)
        QApplication.processEvents()
        
    def output_finished(self, output_path):
        """Handle successful generation of output file"""
        self.progress_bar.setVisible(False)
        
        self.update_output_log(f"Finished generating output file: {output_path}")
        
        # Show success message
        QMessageBox.information(
            self, 
            "Output Generated", 
            f"The merged Excel file has been generated successfully at:\n{output_path}"
        )
        
        # Ask if the user wants to process another file
        reply = QMessageBox.question(
            self, 
            "Process Another?", 
            "Do you want to process another ZIP file?",
            QMessageBox.Yes | QMessageBox.No, 
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.reset_app()
        else:
            self.statusBar().showMessage("Ready")
            
    def output_error(self, error_message):
        """Handle error during output generation"""
        self.progress_bar.setVisible(False)
        self.update_output_log(f"ERROR: {error_message}")
        QMessageBox.critical(self, "Output Error", f"Error generating output file: {error_message}")
        
    def reset_app(self):
        """Reset the application to initial state"""
        # Clear all data
        self.file_data = {}
        self.selected_columns = {}
        self.output_path = None
        
        # Reset UI
        self.file_path_label.clear()
        self.log_label.setText("No processing log yet")
        
        # Clean up temporary directory if it exists
        if hasattr(self, 'temp_dir') and os.path.exists(self.temp_dir):
            try:
                shutil.rmtree(self.temp_dir)
            except Exception as e:
                print(f"Error cleaning up temporary directory: {str(e)}")
                
        # Create a new temporary directory
        self.temp_dir = tempfile.mkdtemp()
        
        # Reset tabs
        self.tabs.setTabEnabled(1, False)
        self.tabs.setTabEnabled(2, False)
        self.tabs.setCurrentIndex(0)
        
        # Reset output path label if it exists
        if hasattr(self, 'output_path_label'):
            self.output_path_label.clear()
            
        # Reset log labels
        if hasattr(self, 'log_label'):
            self.log_label.setText("No processing log yet")
            
        if hasattr(self, 'output_log_label'):
            self.output_log_label.setText("No output log yet")
            
        # Reset summary label
        if hasattr(self, 'summary_label'):
            self.summary_label.setText("No data selected yet")
            
        # Reset progress bar
        self.progress_bar.setVisible(False)
        
        # Show ready status
        self.statusBar().showMessage("Ready")

    def closeEvent(self, event):
        """Clean up on application close"""
        # Clean up temporary directory
        if hasattr(self, 'temp_dir') and os.path.exists(self.temp_dir):
            try:
                shutil.rmtree(self.temp_dir)
            except Exception as e:
                print(f"Error cleaning temporary directory: {e}")
        
        # Save profile settings if available
        if PROFILE_SUPPORT and hasattr(self, 'profile_manager') and self.profile_manager:
            try:
                self.profile_manager.save_settings()
            except Exception as e:
                print(f"Error saving profile settings: {str(e)}")
        
        # Accept the close event
        event.accept()

# Profile management functions
def update_profile_combo(self):
    """Update the profile combo box with available profiles"""
    if not hasattr(self, 'profile_combo') or not PROFILE_SUPPORT:
        return
        
    self.profile_combo.clear()
    
    profiles = self.profile_manager.get_all_profiles()
    default_name = self.profile_manager.default_profile_name
    
    # Add a blank option
    self.profile_combo.addItem("(No profile selected)")
    
    # Add each profile
    for name in sorted(profiles.keys()):
        profile_text = f"{name} (Default)" if name == default_name else name
        self.profile_combo.addItem(profile_text, name)
        
    # Select the default if there is one
    if default_name:
        for i in range(1, self.profile_combo.count()):
            if self.profile_combo.itemData(i) == default_name:
                self.profile_combo.setCurrentIndex(i)
                break

def open_profile_manager(self):
    """Open the profile management dialog"""
    if not PROFILE_SUPPORT or not self.profile_manager:
        QMessageBox.warning(self, "Profile Support Unavailable", 
                           "Profile management is not available in this build.")
        return
        
    dialog = ProfileDialog(
        self, 
        profile_manager=self.profile_manager,
        current_selections=self.selected_columns,
        file_data=self.file_data
    )
    
    # Connect to the profiles_updated signal
    dialog.profiles_updated.connect(self.on_profiles_updated)
    
    result = dialog.exec_()
    
    if result == QDialog.Accepted:
        # Check if a profile was selected and should be applied
        self.on_profiles_updated()
        
        # If a profile was applied and we have file data
        if hasattr(dialog, 'selected_profile') and dialog.selected_profile and self.file_data:
            # Apply the profile selections
            self.apply_profile(dialog.selected_profile)

def on_profiles_updated(self):
    """Handle updates to profiles"""
    if hasattr(self, 'profile_combo'):
        self.update_profile_combo()

def apply_profile(self, profile):
    """Apply a profile to the current data"""
    if not profile or not self.file_data:
        return False
        
    # Get column selections based on the profile
    selections = profile.match_to_new_files(self.file_data)
    
    # Check if we got any selections
    if not selections:
        QMessageBox.information(self, "No Matches", 
                               "The profile did not match any columns in the current files.")
        return False
        
    # Apply the selections
    total_selected = 0
    self.selected_columns = selections
    
    # Count total selections
    for file_name, sheets in self.selected_columns.items():
        for sheet_name, columns in sheets.items():
            total_selected += len(columns)
            
    # Update the UI to reflect selections if we're on the selection tab
    if self.tabs.currentIndex() == 1:
        self.update_checkboxes_for_current_sheet()
        
    QMessageBox.information(self, "Profile Applied", 
                          f"Applied profile '{profile.name}' with {total_selected} columns selected.")
    return True
    
def update_checkboxes_for_current_sheet(self):
    """Update checkboxes for the currently visible sheet based on selections"""
    # Find which sheet is currently displayed
    current_idx = self.sheet_stack.currentIndex()
    current_key = None
    
    for key, idx in self.sheet_widgets.items():
        if idx == current_idx:
            current_key = key
            break
            
    if not current_key:
        return
        
    # Parse the key to get file_name and sheet_name
    parts = current_key.split('_', 1)
    if len(parts) != 2:
        return
        
    file_name, sheet_name = parts
    
    # Update checkboxes for this sheet
    self.update_checkboxes_for_sheet(file_name, sheet_name)

# Add the profile management methods to the ExcelExtractorApp class
ExcelExtractorApp.update_profile_combo = update_profile_combo
ExcelExtractorApp.open_profile_manager = open_profile_manager
ExcelExtractorApp.on_profiles_updated = on_profiles_updated
ExcelExtractorApp.apply_profile = apply_profile
ExcelExtractorApp.update_checkboxes_for_current_sheet = update_checkboxes_for_current_sheet

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
        
        # Enable native macOS dark mode support
        os.environ['QT_MAC_WANTS_LAYER'] = '1'
    
    app = QApplication(sys.argv)
    
    # Apply styling based on platform
    if sys.platform == 'darwin':
        # Use macOS native style for best integration including dark mode support
        app.setStyle("macintosh")
        
        # Enable automatic palette adjustment based on system appearance
        app.setAttribute(Qt.AA_DontCreateNativeWidgetSiblings)
    else:
        # For non-macOS platforms, use Fusion style which looks modern
        app.setStyle("Fusion")
    
    # Set application-wide attribute to use the native color scheme (respects dark mode)
    app.setAttribute(Qt.AA_UseStyleSheetPropagationInWidgetStyles, True)
    
    window = ExcelExtractorApp()
    window.show()
    
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()