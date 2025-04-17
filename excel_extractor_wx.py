#!/usr/bin/env python3
"""
Excel Data Extractor - wxPython MacOS Application
This application extracts and merges selected data from multiple Excel files in a ZIP archive.
Optimized exclusively for MacOS.
"""

import os
import sys
import tempfile
import pandas as pd
from zipfile import ZipFile
from pathlib import Path
import xlwt
import wx
import wx.grid
import wx.lib.scrolledpanel as scrolled
import wx.lib.agw.multidirdialog as MDD
import threading

# Force light mode for the application (needed for dark mode macOS)
os.environ['PYOPENGL_PLATFORM'] = 'egl'  # Prevent dark mode issues with OpenGL

# Constants
APP_NAME = "Excel Data Extractor"
APP_VERSION = "1.0.0"

class ExcelExtractorFrame(wx.Frame):
    def __init__(self):
        # Initialize the parent class
        wx.Frame.__init__(self, None, title=APP_NAME, size=(900, 600))
        
        # Set the application icon
        self.SetIcon(wx.Icon(wx.ArtProvider.GetBitmap(wx.ART_NORMAL_FILE, wx.ART_FRAME_ICON)))
        
        # Initialize instance variables
        self.file_data = {}
        self.selected_columns = {}
        self.temp_dir = None
        self.output_path = None
        
        # Create the UI
        self.create_ui()
        
        # Bind the close event
        self.Bind(wx.EVT_CLOSE, self.on_close)
        
        # Center the window on screen
        self.Center()
        
        # Set MacOS-specific features
        self.configure_for_macos()
    
    def configure_for_macos(self):
        """Configure the application specifically for MacOS"""
        if 'darwin' in sys.platform:
            # Create menu bar for MacOS
            menubar = wx.MenuBar()
            file_menu = wx.Menu()
            
            # Add menu items
            open_item = file_menu.Append(wx.ID_OPEN, "&Open ZIP File...\tCtrl+O", "Open a ZIP file")
            file_menu.AppendSeparator()
            exit_item = file_menu.Append(wx.ID_EXIT, "E&xit", "Exit the application")
            
            # Add the file menu to the menu bar
            menubar.Append(file_menu, "&File")
            
            # Set the menu bar
            self.SetMenuBar(menubar)
            
            # Bind menu events
            self.Bind(wx.EVT_MENU, self.on_open, open_item)
            self.Bind(wx.EVT_MENU, self.on_exit, exit_item)
    
    def create_ui(self):
        """Create the main user interface"""
        # Create a panel for the UI with light background (for macOS dark mode compatibility)
        self.panel = wx.Panel(self)
        self.panel.SetBackgroundColour(wx.WHITE)
        
        # Create a notebook for the tabs with visible styling
        self.notebook = wx.Notebook(self.panel)
        
        # Create the tabs with light background
        self.upload_tab = wx.Panel(self.notebook)
        self.upload_tab.SetBackgroundColour(wx.WHITE)
        
        self.selection_tab = wx.Panel(self.notebook)
        self.selection_tab.SetBackgroundColour(wx.WHITE)
        
        self.output_tab = wx.Panel(self.notebook)
        self.output_tab.SetBackgroundColour(wx.WHITE)
        
        # Add the tabs to the notebook with clear labels
        self.notebook.AddPage(self.upload_tab, "1. Upload ZIP")
        self.notebook.AddPage(self.selection_tab, "2. Select Data")
        self.notebook.AddPage(self.output_tab, "3. Generate Output")
        
        # Disable tabs until they're ready (compatible with all wxPython versions)
        self.notebook.SetSelection(0)  # Start with the first tab
        # We'll handle tab enabling/disabling in the code by changing pages programmatically
        
        # Create the UI for each tab
        self.create_upload_tab()
        self.create_selection_tab()
        self.create_output_tab()
        
        # Create a sizer for the main panel
        panel_sizer = wx.BoxSizer(wx.VERTICAL)
        
        # Add a heading
        heading = wx.StaticText(self.panel, label=APP_NAME)
        heading_font = heading.GetFont()
        heading_font.SetPointSize(16)
        heading_font.SetWeight(wx.FONTWEIGHT_BOLD)
        heading.SetFont(heading_font)
        heading.SetForegroundColour(wx.BLACK)  # Ensure visible text in dark mode
        
        # Add a description
        description = wx.StaticText(self.panel, label="Extract and merge data from multiple Excel files in a ZIP archive")
        description.SetForegroundColour(wx.BLACK)  # Ensure visible text in dark mode
        
        # Add elements to the sizer
        panel_sizer.Add(heading, 0, wx.ALIGN_CENTER | wx.TOP, 10)
        panel_sizer.Add(description, 0, wx.ALIGN_CENTER | wx.BOTTOM, 10)
        panel_sizer.Add(self.notebook, 1, wx.EXPAND | wx.ALL, 10)
        
        # Create a status bar
        self.status_bar = self.CreateStatusBar()
        self.status_bar.SetStatusText("Ready")
        
        # Set the sizer for the panel
        self.panel.SetSizer(panel_sizer)
    
    def create_upload_tab(self):
        """Create the UI for the upload tab"""
        # Create a sizer for the upload tab
        upload_sizer = wx.BoxSizer(wx.VERTICAL)
        
        # Add instructions
        instructions = wx.StaticText(self.upload_tab, label="Select a ZIP file containing Excel files (.xlsx or .xls).\nThe application will extract the Excel files and let you choose which data to extract.")
        upload_sizer.Add(instructions, 0, wx.ALL, 10)
        
        # Add MacOS tips
        tips_box = wx.StaticBox(self.upload_tab, label="MacOS Tips")
        tips_sizer = wx.StaticBoxSizer(tips_box, wx.VERTICAL)
        tips = wx.StaticText(tips_box, label="• Create a ZIP file by selecting multiple Excel files, right-clicking, and choosing 'Compress'\n• Make sure your Excel files are readable and not password-protected\n• Avoid using special characters in filenames")
        tips_sizer.Add(tips, 0, wx.ALL, 10)
        upload_sizer.Add(tips_sizer, 0, wx.EXPAND | wx.ALL, 10)
        
        # Create a file picker
        file_sizer = wx.BoxSizer(wx.HORIZONTAL)
        file_label = wx.StaticText(self.upload_tab, label="ZIP File:")
        self.file_picker = wx.FilePickerCtrl(
            self.upload_tab,
            wildcard="ZIP files (*.zip)|*.zip",
            style=wx.FLP_OPEN | wx.FLP_FILE_MUST_EXIST | wx.FLP_USE_TEXTCTRL
        )
        
        file_sizer.Add(file_label, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        file_sizer.Add(self.file_picker, 1, wx.EXPAND)
        upload_sizer.Add(file_sizer, 0, wx.EXPAND | wx.ALL, 10)
        
        # Add a process button
        self.process_button = wx.Button(self.upload_tab, label="Process ZIP File")
        self.process_button.Bind(wx.EVT_BUTTON, self.on_process_zip)
        upload_sizer.Add(self.process_button, 0, wx.ALIGN_CENTER | wx.ALL, 10)
        
        # Add a log area
        log_box = wx.StaticBox(self.upload_tab, label="Processing Log")
        log_sizer = wx.StaticBoxSizer(log_box, wx.VERTICAL)
        self.log_text = wx.TextCtrl(log_box, style=wx.TE_MULTILINE | wx.TE_READONLY)
        log_sizer.Add(self.log_text, 1, wx.EXPAND | wx.ALL, 5)
        upload_sizer.Add(log_sizer, 1, wx.EXPAND | wx.ALL, 10)
        
        # Set the sizer for the upload tab
        self.upload_tab.SetSizer(upload_sizer)
    
    def create_selection_tab(self):
        """Create the UI for the selection tab"""
        # Create a sizer for the selection tab
        selection_sizer = wx.BoxSizer(wx.VERTICAL)
        
        # Add a placeholder
        self.selection_placeholder = wx.StaticText(
            self.selection_tab, 
            label="Data selection will be available after processing a ZIP file"
        )
        self.selection_placeholder.SetForegroundColour(wx.BLACK)  # Ensure visible text in dark mode
        selection_sizer.Add(self.selection_placeholder, 1, wx.ALIGN_CENTER)
        
        # The actual selection UI will be created dynamically after processing a ZIP file
        
        # Set the sizer for the selection tab
        self.selection_tab.SetSizer(selection_sizer)
    
    def create_output_tab(self):
        """Create the UI for the output tab"""
        # Create a sizer for the output tab
        output_sizer = wx.BoxSizer(wx.VERTICAL)
        
        # Add instructions
        instructions = wx.StaticText(
            self.output_tab, 
            label="Enter a name for the output Excel file and select where to save it.\nClick 'Process and Generate' to create the merged Excel file."
        )
        instructions.SetForegroundColour(wx.BLACK)  # Ensure visible text in dark mode
        output_sizer.Add(instructions, 0, wx.ALL, 10)
        
        # Add filename input
        name_sizer = wx.BoxSizer(wx.HORIZONTAL)
        name_label = wx.StaticText(self.output_tab, label="Output filename:")
        name_label.SetForegroundColour(wx.BLACK)  # Ensure visible text in dark mode
        self.output_name = wx.TextCtrl(self.output_tab, value="merged_data")
        name_sizer.Add(name_label, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        name_sizer.Add(self.output_name, 1, wx.EXPAND)
        output_sizer.Add(name_sizer, 0, wx.EXPAND | wx.ALL, 10)
        
        # Add location picker
        location_sizer = wx.BoxSizer(wx.HORIZONTAL)
        location_label = wx.StaticText(self.output_tab, label="Save location:")
        location_label.SetForegroundColour(wx.BLACK)  # Ensure visible text in dark mode
        self.location_picker = wx.DirPickerCtrl(
            self.output_tab,
            style=wx.DIRP_USE_TEXTCTRL | wx.DIRP_DIR_MUST_EXIST
        )
        location_sizer.Add(location_label, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        location_sizer.Add(self.location_picker, 1, wx.EXPAND)
        output_sizer.Add(location_sizer, 0, wx.EXPAND | wx.ALL, 10)
        
        # Add a process button
        self.generate_button = wx.Button(self.output_tab, label="Process and Generate Excel File")
        self.generate_button.Bind(wx.EVT_BUTTON, self.on_generate_output)
        output_sizer.Add(self.generate_button, 0, wx.ALIGN_CENTER | wx.ALL, 10)
        
        # Add a log area
        output_log_box = wx.StaticBox(self.output_tab, label="Processing Log")
        output_log_sizer = wx.StaticBoxSizer(output_log_box, wx.VERTICAL)
        self.output_log_text = wx.TextCtrl(output_log_box, style=wx.TE_MULTILINE | wx.TE_READONLY)
        output_log_sizer.Add(self.output_log_text, 1, wx.EXPAND | wx.ALL, 5)
        output_sizer.Add(output_log_sizer, 1, wx.EXPAND | wx.ALL, 10)
        
        # Add navigation buttons
        nav_sizer = wx.BoxSizer(wx.HORIZONTAL)
        back_button = wx.Button(self.output_tab, label="Back to Selection")
        back_button.Bind(wx.EVT_BUTTON, lambda event: self.notebook.SetSelection(1))
        nav_sizer.Add(back_button, 0, wx.RIGHT, 10)
        output_sizer.Add(nav_sizer, 0, wx.ALIGN_LEFT | wx.ALL, 10)
        
        # Set the sizer for the output tab
        self.output_tab.SetSizer(output_sizer)
    
    def create_dynamic_selection_ui(self):
        """Create the dynamic UI for data selection based on the loaded files"""
        # Clear the selection tab
        self.selection_tab.DestroyChildren()
        
        # Create a new sizer
        selection_sizer = wx.BoxSizer(wx.VERTICAL)
        
        # Create a notebook for the files
        file_notebook = wx.Notebook(self.selection_tab)
        
        # Create a tab for each Excel file
        for file_name, sheets in self.file_data.items():
            # Create a panel for this file
            file_panel = wx.Panel(file_notebook)
            file_sizer = wx.BoxSizer(wx.VERTICAL)
            
            # Create a notebook for the sheets in this file
            sheet_notebook = wx.Notebook(file_panel)
            
            # Create a tab for each sheet
            for sheet_name, df in sheets.items():
                # Create a panel for this sheet
                sheet_panel = wx.Panel(sheet_notebook)
                sheet_sizer = wx.BoxSizer(wx.VERTICAL)
                
                # Add a preview of the data
                preview_box = wx.StaticBox(sheet_panel, label="Data Preview")
                preview_sizer = wx.StaticBoxSizer(preview_box, wx.VERTICAL)
                
                # Create a grid for the data preview
                preview_grid = wx.grid.Grid(preview_box)
                
                # Get the first 5 rows for preview
                preview_df = df.head(5)
                
                # Set up the grid
                preview_grid.CreateGrid(len(preview_df), len(preview_df.columns))
                
                # Set column headers
                for col_idx, col_name in enumerate(preview_df.columns):
                    preview_grid.SetColLabelValue(col_idx, str(col_name))
                
                # Fill in the data
                for row_idx, (_, row) in enumerate(preview_df.iterrows()):
                    for col_idx, value in enumerate(row):
                        # Handle NaN values
                        if pd.isna(value):
                            cell_value = ""
                        else:
                            cell_value = str(value)
                        preview_grid.SetCellValue(row_idx, col_idx, cell_value)
                
                # Auto-size columns
                preview_grid.AutoSizeColumns()
                
                preview_sizer.Add(preview_grid, 1, wx.EXPAND | wx.ALL, 5)
                sheet_sizer.Add(preview_sizer, 1, wx.EXPAND | wx.ALL, 10)
                
                # Add column selection
                selection_box = wx.StaticBox(sheet_panel, label="Select Columns to Extract")
                selection_sizer = wx.StaticBoxSizer(selection_box, wx.VERTICAL)
                
                # Add select all/none buttons
                button_sizer = wx.BoxSizer(wx.HORIZONTAL)
                select_all_button = wx.Button(selection_box, label="Select All")
                deselect_all_button = wx.Button(selection_box, label="Deselect All")
                
                # Bind button events with specific file and sheet names
                select_all_button.file_name = file_name
                select_all_button.sheet_name = sheet_name
                deselect_all_button.file_name = file_name
                deselect_all_button.sheet_name = sheet_name
                
                select_all_button.Bind(wx.EVT_BUTTON, self.on_select_all)
                deselect_all_button.Bind(wx.EVT_BUTTON, self.on_deselect_all)
                
                button_sizer.Add(select_all_button, 0, wx.RIGHT, 10)
                button_sizer.Add(deselect_all_button, 0)
                selection_sizer.Add(button_sizer, 0, wx.ALL, 5)
                
                # Create a scrolled panel for the checkboxes
                checkbox_panel = scrolled.ScrolledPanel(selection_box)
                checkbox_sizer = wx.FlexGridSizer(0, 3, 5, 20)  # rows, cols, vgap, hgap
                
                # Make sure the selected_columns structure is initialized
                if file_name not in self.selected_columns:
                    self.selected_columns[file_name] = {}
                if sheet_name not in self.selected_columns[file_name]:
                    self.selected_columns[file_name][sheet_name] = []
                
                # Create a checkbox for each column
                for col_name in df.columns:
                    checkbox = wx.CheckBox(checkbox_panel, label=str(col_name))
                    
                    # Store column info in the checkbox
                    checkbox.file_name = file_name
                    checkbox.sheet_name = sheet_name
                    checkbox.column_name = col_name
                    
                    # Bind the checkbox event
                    checkbox.Bind(wx.EVT_CHECKBOX, self.on_column_checkbox)
                    
                    # Add to the sizer
                    checkbox_sizer.Add(checkbox, 0, wx.ALL, 2)
                
                # Set up the scrolled panel
                checkbox_panel.SetSizer(checkbox_sizer)
                checkbox_panel.SetupScrolling()
                checkbox_panel.SetMinSize((-1, 150))  # Set a minimum height
                
                selection_sizer.Add(checkbox_panel, 1, wx.EXPAND | wx.ALL, 5)
                sheet_sizer.Add(selection_sizer, 0, wx.EXPAND | wx.ALL, 10)
                
                # Set the sizer for the sheet panel
                sheet_panel.SetSizer(sheet_sizer)
                
                # Add the sheet tab to the notebook
                sheet_notebook.AddPage(sheet_panel, sheet_name)
            
            # Add the sheet notebook to the file panel
            file_sizer.Add(sheet_notebook, 1, wx.EXPAND | wx.ALL, 10)
            
            # Add navigation buttons
            nav_sizer = wx.BoxSizer(wx.HORIZONTAL)
            back_button = wx.Button(file_panel, label="Back to Upload")
            next_button = wx.Button(file_panel, label="Continue to Output")
            
            # Bind button events
            back_button.Bind(wx.EVT_BUTTON, lambda event: self.notebook.SetSelection(0))
            next_button.Bind(wx.EVT_BUTTON, self.on_continue_to_output)
            
            nav_sizer.Add(back_button, 0, wx.RIGHT, 10)
            nav_sizer.Add(next_button, 0)
            file_sizer.Add(nav_sizer, 0, wx.ALIGN_RIGHT | wx.ALL, 10)
            
            # Set the sizer for the file panel
            file_panel.SetSizer(file_sizer)
            
            # Add the file tab to the notebook
            file_notebook.AddPage(file_panel, file_name)
        
        # Add the file notebook to the selection tab
        selection_sizer.Add(file_notebook, 1, wx.EXPAND | wx.ALL, 10)
        
        # Add a status line showing total selected columns
        self.selection_status = wx.StaticText(self.selection_tab, label="Total columns selected: 0")
        self.selection_status.SetForegroundColour(wx.BLACK)  # Ensure visible text in dark mode
        selection_sizer.Add(self.selection_status, 0, wx.ALL, 10)
        
        # Set the sizer for the selection tab
        self.selection_tab.SetSizer(selection_sizer)
        
        # Update the layout
        self.selection_tab.Layout()
    
    def update_log(self, message):
        """Update the log with a new message"""
        # Use CallAfter to update the UI from a different thread
        wx.CallAfter(self.log_text.AppendText, message + "\n")
    
    def update_output_log(self, message):
        """Update the output log with a new message"""
        # Use CallAfter to update the UI from a different thread
        wx.CallAfter(self.output_log_text.AppendText, message + "\n")
    
    def update_status(self, message):
        """Update the status bar message"""
        # Use CallAfter to update the UI from a different thread
        wx.CallAfter(self.status_bar.SetStatusText, message)
    
    def update_selection_status(self):
        """Update the selection status showing total selected columns"""
        # Calculate total selected columns
        total_selected = sum(len(cols) for file in self.selected_columns.values() for cols in file.values())
        
        # Update the status text
        if hasattr(self, 'selection_status'):
            self.selection_status.SetLabel(f"Total columns selected: {total_selected}")
            
            # Update the status bar too
            self.status_bar.SetStatusText(f"Selected {total_selected} columns")
    
    def on_open(self, event):
        """Handle the File -> Open menu event"""
        # Create a file dialog
        with wx.FileDialog(
            self, "Open ZIP File", wildcard="ZIP files (*.zip)|*.zip",
            style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST
        ) as fileDialog:
            
            # If the user clicked OK
            if fileDialog.ShowModal() == wx.ID_OK:
                # Get the path
                self.file_picker.SetPath(fileDialog.GetPath())
    
    def on_exit(self, event):
        """Handle the File -> Exit menu event"""
        self.Close()
    
    def on_process_zip(self, event):
        """Handle the Process ZIP button event"""
        # Get the ZIP file path
        zip_path = self.file_picker.GetPath()
        
        # Check if a file was selected
        if not zip_path:
            wx.MessageBox(
                "Please select a ZIP file containing Excel files.",
                "No File Selected",
                wx.OK | wx.ICON_WARNING
            )
            return
        
        # Check if the file exists
        if not os.path.exists(zip_path):
            wx.MessageBox(
                f"The selected file does not exist: {zip_path}",
                "File Not Found",
                wx.OK | wx.ICON_WARNING
            )
            return
        
        # Clear the log
        self.log_text.Clear()
        
        # Disable the process button and file picker
        self.process_button.Disable()
        self.file_picker.Disable()
        
        # Create a temporary directory for extraction
        self.temp_dir = tempfile.mkdtemp()
        
        # Update the status
        self.update_status("Processing ZIP file...")
        
        # Create a thread for processing
        threading.Thread(target=self.process_zip_thread, args=(zip_path,)).start()
    
    def process_zip_thread(self, zip_path):
        """Process the ZIP file in a separate thread"""
        try:
            # Extract Excel files from the ZIP
            self.update_log("Extracting Excel files from ZIP...")
            extracted_files = self.extract_zip_file(zip_path, self.temp_dir)
            
            if not extracted_files:
                wx.CallAfter(wx.MessageBox,
                    "No Excel files found in the ZIP archive.",
                    "No Excel Files",
                    wx.OK | wx.ICON_WARNING
                )
                # Re-enable the UI
                wx.CallAfter(self.process_button.Enable)
                wx.CallAfter(self.file_picker.Enable)
                wx.CallAfter(self.update_status, "No Excel files found")
                return
            
            # Read Excel files and store their data
            self.update_log("Reading Excel files...")
            self.file_data = self.read_excel_files(extracted_files)
            
            if not self.file_data:
                wx.CallAfter(wx.MessageBox,
                    "Could not read any data from the Excel files.",
                    "No Data",
                    wx.OK | wx.ICON_WARNING
                )
                # Re-enable the UI
                wx.CallAfter(self.process_button.Enable)
                wx.CallAfter(self.file_picker.Enable)
                wx.CallAfter(self.update_status, "No data in Excel files")
                return
            
            # Initialize the selected_columns structure
            self.selected_columns = {}
            for file_name, sheets in self.file_data.items():
                self.selected_columns[file_name] = {}
                for sheet_name in sheets.keys():
                    self.selected_columns[file_name][sheet_name] = []
            
            # Create the dynamic selection UI
            wx.CallAfter(self.create_dynamic_selection_ui)
            
            # Switch to the selection tab (no need to explicitly enable in newer wxPython)
            wx.CallAfter(self.notebook.SetSelection, 1)
            
            # Update the status
            wx.CallAfter(self.update_status, "Ready to select data")
            
        except Exception as e:
            # Show error message
            wx.CallAfter(wx.MessageBox,
                f"An error occurred while processing the ZIP file:\n{str(e)}",
                "Processing Error",
                wx.OK | wx.ICON_ERROR
            )
            
            # Log the error
            self.update_log(f"Error: {str(e)}")
            
            # Update the status
            wx.CallAfter(self.update_status, "Error processing ZIP file")
            
        finally:
            # Re-enable the UI
            wx.CallAfter(self.process_button.Enable)
            wx.CallAfter(self.file_picker.Enable)
    
    def extract_zip_file(self, zip_path, extract_dir):
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
            self.update_log(f"Opening ZIP file: {zip_path}")
            
            with ZipFile(zip_path, 'r') as zip_ref:
                # List all files in the ZIP
                file_list = zip_ref.namelist()
                
                self.update_log(f"Found {len(file_list)} files in ZIP archive")
                
                # Extract only Excel files
                for file_name in file_list:
                    lower_name = file_name.lower()
                    if lower_name.endswith('.xlsx') or lower_name.endswith('.xls'):
                        # Handle folder paths in ZIP
                        if file_name.endswith('/') or os.path.basename(file_name) == '':
                            continue
                            
                        # Extract the file
                        try:
                            self.update_log(f"Extracting: {file_name}")
                            zip_ref.extract(file_name, extract_dir)
                            full_path = os.path.join(extract_dir, file_name)
                            excel_files.append(full_path)
                        except Exception as extract_error:
                            self.update_log(f"Could not extract {file_name}: {str(extract_error)}")
                    
                # Also look for Excel files in any folders that were extracted
                for root, dirs, files in os.walk(extract_dir):
                    for file in files:
                        if file.lower().endswith(('.xlsx', '.xls')) and os.path.join(root, file) not in excel_files:
                            excel_files.append(os.path.join(root, file))
                            self.update_log(f"Found additional Excel file: {file}")
        
        except Exception as e:
            self.update_log(f"Error extracting ZIP file: {str(e)}")
            return []
        
        self.update_log(f"Extracted {len(excel_files)} Excel files")
        return excel_files
    
    def read_excel_files(self, file_paths):
        """
        Read data from multiple Excel files
        
        Parameters:
        - file_paths: List of paths to Excel files
        
        Returns:
        - A nested dictionary structure: {file_name: {sheet_name: dataframe}}
        """
        file_data = {}
        
        if not file_paths:
            self.update_log("No Excel files to process")
            return file_data
        
        self.update_log(f"Reading {len(file_paths)} Excel files...")
        
        for file_path in file_paths:
            try:
                # Get just the filename without path
                file_name = os.path.basename(file_path)
                self.update_log(f"Reading: {file_name}")
                
                # Read all sheets from the Excel file
                try:
                    excel_file = pd.ExcelFile(file_path)
                    sheet_names = excel_file.sheet_names
                    self.update_log(f"Found {len(sheet_names)} sheets in {file_name}")
                except Exception as excel_error:
                    self.update_log(f"Error opening Excel file '{file_name}': {str(excel_error)}")
                    
                    # Try alternate approach for older Excel formats
                    try:
                        # For xls files
                        if file_path.lower().endswith('.xls'):
                            df = pd.read_excel(file_path, engine='xlrd')
                            file_data[file_name] = {"Sheet1": df}
                            self.update_log(f"Successfully read {file_name} using xlrd engine")
                            continue
                    except Exception as alt_error:
                        self.update_log(f"Alternative read approach failed: {str(alt_error)}")
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
                            self.update_log(f"Sheet '{sheet_name}' has {len(df)} rows and {len(df.columns)} columns")
                        else:
                            self.update_log(f"Sheet '{sheet_name}' is empty, skipping")
                    except Exception as e:
                        self.update_log(f"Error reading sheet '{sheet_name}': {str(e)}")
                        continue
                
                # If no sheets were successfully read, remove this file entry
                if not file_data[file_name]:
                    self.update_log(f"No data found in file '{file_name}'")
                    del file_data[file_name]
                    
            except Exception as e:
                self.update_log(f"Error reading file '{os.path.basename(file_path)}': {str(e)}")
                continue
        
        # Provide summary
        file_count = len(file_data)
        if file_count > 0:
            sheet_count = sum(len(sheets) for sheets in file_data.values())
            self.update_log(f"Successfully read {file_count} files with a total of {sheet_count} sheets")
        else:
            self.update_log("Could not read any data from the Excel files")
        
        return file_data
    
    def on_column_checkbox(self, event):
        """Handle column selection checkbox events"""
        # Get the checkbox
        checkbox = event.GetEventObject()
        
        # Get the column info
        file_name = checkbox.file_name
        sheet_name = checkbox.sheet_name
        column_name = checkbox.column_name
        
        # Update the selected columns structure
        if checkbox.IsChecked():
            if column_name not in self.selected_columns[file_name][sheet_name]:
                self.selected_columns[file_name][sheet_name].append(column_name)
        else:
            if column_name in self.selected_columns[file_name][sheet_name]:
                self.selected_columns[file_name][sheet_name].remove(column_name)
        
        # Update the selection status
        self.update_selection_status()
    
    def on_select_all(self, event):
        """Handle select all button events"""
        # Get the button
        button = event.GetEventObject()
        
        # Get the file and sheet info
        file_name = button.file_name
        sheet_name = button.sheet_name
        
        # Get all columns for this sheet
        all_columns = self.file_data[file_name][sheet_name].columns.tolist()
        
        # Update the selected columns structure
        self.selected_columns[file_name][sheet_name] = all_columns.copy()
        
        # Update all checkboxes for this sheet
        self.update_sheet_checkboxes(file_name, sheet_name)
        
        # Update the selection status
        self.update_selection_status()
    
    def on_deselect_all(self, event):
        """Handle deselect all button events"""
        # Get the button
        button = event.GetEventObject()
        
        # Get the file and sheet info
        file_name = button.file_name
        sheet_name = button.sheet_name
        
        # Clear the selected columns for this sheet
        self.selected_columns[file_name][sheet_name] = []
        
        # Update all checkboxes for this sheet
        self.update_sheet_checkboxes(file_name, sheet_name)
        
        # Update the selection status
        self.update_selection_status()
    
    def update_sheet_checkboxes(self, file_name, sheet_name):
        """Update all checkboxes for a specific sheet"""
        # Find all checkboxes for this sheet and update them
        for window in wx.GetTopLevelWindows():
            if isinstance(window, wx.Frame) and window == self:
                # Recursively find all checkboxes
                for checkbox in self.find_checkboxes(window):
                    if (hasattr(checkbox, 'file_name') and 
                        hasattr(checkbox, 'sheet_name') and 
                        hasattr(checkbox, 'column_name')):
                        if (checkbox.file_name == file_name and 
                            checkbox.sheet_name == sheet_name):
                            # Set the checkbox state without triggering events
                            checkbox.SetValue(
                                checkbox.column_name in self.selected_columns[file_name][sheet_name]
                            )
    
    def find_checkboxes(self, parent):
        """Recursively find all checkboxes in a parent window"""
        checkboxes = []
        for child in parent.GetChildren():
            if isinstance(child, wx.CheckBox):
                checkboxes.append(child)
            checkboxes.extend(self.find_checkboxes(child))
        return checkboxes
    
    def on_continue_to_output(self, event):
        """Handle the Continue to Output button event"""
        # Check if any columns are selected
        total_selected = sum(len(cols) for file in self.selected_columns.values() for cols in file.values())
        
        if total_selected == 0:
            wx.MessageBox(
                "Please select at least one column to extract.",
                "No Columns Selected",
                wx.OK | wx.ICON_WARNING
            )
            return
        
        # Switch to the output tab (no need to explicitly enable in newer wxPython)
        self.notebook.SetSelection(2)
    
    def on_generate_output(self, event):
        """Handle the Process and Generate button event"""
        # Get the output filename
        output_name = self.output_name.GetValue().strip()
        
        # Check if a name was entered
        if not output_name:
            wx.MessageBox(
                "Please enter a name for the output file.",
                "No Filename",
                wx.OK | wx.ICON_WARNING
            )
            return
        
        # Add extension if not present
        if not output_name.lower().endswith('.xls'):
            output_name += ".xls"
        
        # Get the output location
        output_dir = self.location_picker.GetPath()
        
        # Check if a location was selected
        if not output_dir:
            # Use file dialog to get a save location
            with wx.FileDialog(
                self, "Save Output Excel File", 
                wildcard="Excel files (*.xls)|*.xls",
                defaultFile=output_name,
                style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT
            ) as fileDialog:
                
                # If the user clicked OK
                if fileDialog.ShowModal() == wx.ID_OK:
                    output_path = fileDialog.GetPath()
                else:
                    return  # User cancelled
        else:
            # Use the selected directory
            output_path = os.path.join(output_dir, output_name)
        
        # Store the output path
        self.output_path = output_path
        
        # Clear the output log
        self.output_log_text.Clear()
        
        # Disable the generate button
        self.generate_button.Disable()
        
        # Update the status
        self.update_status("Generating output file...")
        
        # Create a thread for processing
        threading.Thread(target=self.generate_output_thread).start()
    
    def generate_output_thread(self):
        """Generate the output Excel file in a separate thread"""
        try:
            # Process and merge the selected data
            self.update_output_log(f"Generating output file: {self.output_path}")
            success = self.process_and_merge_data()
            
            if success:
                # Show success message
                wx.CallAfter(wx.MessageBox,
                    f"The merged Excel file has been saved to:\n{self.output_path}",
                    "Processing Complete",
                    wx.OK | wx.ICON_INFORMATION
                )
                
                # Update the status
                wx.CallAfter(self.update_status, "Processing complete")
                
                # Ask if user wants to process another file
                wx.CallAfter(self.ask_process_another)
            else:
                # Show error message
                wx.CallAfter(wx.MessageBox,
                    "Failed to process and merge data.",
                    "Processing Error",
                    wx.OK | wx.ICON_ERROR
                )
                
                # Update the status
                wx.CallAfter(self.update_status, "Processing failed")
        
        except Exception as e:
            # Show error message
            wx.CallAfter(wx.MessageBox,
                f"An error occurred while generating the output file:\n{str(e)}",
                "Processing Error",
                wx.OK | wx.ICON_ERROR
            )
            
            # Log the error
            self.update_output_log(f"Error: {str(e)}")
            
            # Update the status
            wx.CallAfter(self.update_status, "Error generating output file")
        
        finally:
            # Re-enable the generate button
            wx.CallAfter(self.generate_button.Enable)
    
    def process_and_merge_data(self):
        """
        Process and merge selected data from multiple Excel files
        
        Returns:
        - True if successful, False otherwise
        """
        try:
            self.update_output_log("Starting data processing...")
            
            # Create a new workbook
            workbook = xlwt.Workbook()
            
            # Track the number of worksheets created
            worksheet_count = 0
            
            # Process each file
            for file_name, sheets in self.file_data.items():
                self.update_output_log(f"Processing file: {file_name}")
                
                # Process each sheet in the file
                for sheet_name, df in sheets.items():
                    # Get the selected columns for this sheet
                    cols = self.selected_columns.get(file_name, {}).get(sheet_name, [])
                    
                    # Skip if no columns were selected for this sheet
                    if not cols:
                        self.update_output_log(f"No columns selected for {file_name} - {sheet_name}, skipping")
                        continue
                    
                    self.update_output_log(f"Processing sheet: {sheet_name} with {len(cols)} selected columns")
                    
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
            self.update_output_log(f"Saving output to: {self.output_path}")
            workbook.save(self.output_path)
            
            self.update_output_log(f"Processing complete. Created {worksheet_count} worksheets plus summary.")
            return True
        
        except Exception as e:
            self.update_output_log(f"Error processing and merging data: {str(e)}")
            return False
    
    def ask_process_another(self):
        """Ask if the user wants to process another file"""
        dlg = wx.MessageDialog(
            self,
            "Would you like to process another ZIP file?",
            "Process Another?",
            wx.YES_NO | wx.ICON_QUESTION
        )
        
        if dlg.ShowModal() == wx.ID_YES:
            self.reset_app()
        
        dlg.Destroy()
    
    def reset_app(self):
        """Reset the application to initial state"""
        # Clear all data
        self.file_data = {}
        self.selected_columns = {}
        self.output_path = None
        
        # Clear UI elements
        self.file_picker.SetPath("")
        self.log_text.Clear()
        self.output_log_text.Clear()
        self.location_picker.SetPath("")
        
        # Reset to the first tab (upload)
        self.notebook.SetSelection(0)
        
        # Clean up temporary directory
        if self.temp_dir and os.path.exists(self.temp_dir):
            try:
                import shutil
                shutil.rmtree(self.temp_dir)
                self.temp_dir = None
            except Exception as e:
                print(f"Error cleaning temporary directory: {e}")
        
        # Update status
        self.status_bar.SetStatusText("Ready")
    
    def on_close(self, event):
        """Handle the window close event"""
        # Clean up temporary directory
        if self.temp_dir and os.path.exists(self.temp_dir):
            try:
                import shutil
                shutil.rmtree(self.temp_dir)
            except Exception as e:
                print(f"Error cleaning temporary directory: {e}")
        
        # Destroy the window
        self.Destroy()

class ExcelExtractorApp(wx.App):
    def OnInit(self):
        # Create the main frame
        frame = ExcelExtractorFrame()
        frame.Show()
        
        # Set the frame as the top window
        self.SetTopWindow(frame)
        
        return True

def main():
    # Create the application
    app = ExcelExtractorApp(False)
    
    # Start the event loop
    app.MainLoop()

if __name__ == "__main__":
    main()