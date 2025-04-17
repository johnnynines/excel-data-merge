"""
Excel Data Extractor - Profile Dialog
This module provides the UI components for managing extraction profiles.
"""

import os
import re
from typing import Dict, List, Optional, Any, Tuple, Set
import json

from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QListWidget, 
    QListWidgetItem, QTabWidget, QWidget, QLineEdit, QFormLayout, QCheckBox,
    QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox, QFileDialog,
    QGroupBox, QComboBox
)
from PyQt5.QtCore import Qt, QSize, pyqtSignal
from PyQt5.QtGui import QIcon, QFont, QColor

# Import profile management (with typing-only import for type hints)
from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from profile_manager import ProfileManager, ExtractionProfile


class ProfileListItem(QListWidgetItem):
    """Custom list item for displaying profiles"""
    
    def __init__(self, profile: 'ExtractionProfile', is_default: bool = False):
        """Initialize a profile list item"""
        # Create display text (mark default profile)
        display_text = f"{profile.name} {'(Default)' if is_default else ''}"
        
        # Call parent constructor
        super().__init__(display_text)
        
        # Store reference to the profile
        self.profile = profile
        self.is_default = is_default


class ProfileDialog(QDialog):
    """Dialog for managing extraction profiles"""
    
    # Signal emitted when profiles are updated
    profiles_updated = pyqtSignal()
    
    def __init__(self, parent=None, profile_manager: Optional['ProfileManager'] = None, 
                 current_selections: Optional[Dict[str, Dict[str, List[str]]]] = None,
                 file_data: Optional[Dict[str, Dict[str, Any]]] = None):
        """Initialize the profile dialog"""
        super().__init__(parent)
        
        # Store parameters
        self.profile_manager = profile_manager
        self.current_selections = current_selections or {}
        self.file_data = file_data or {}
        self.selected_profile = None
        
        # Initialize UI
        self.setup_ui()
        
    def setup_ui(self):
        """Setup the dialog UI"""
        # Set dialog properties
        self.setWindowTitle("Manage Extraction Profiles")
        self.setMinimumSize(700, 500)
        
        # Create main layout
        main_layout = QHBoxLayout(self)
        
        # Create profile list panel (left side)
        list_panel = QWidget()
        list_layout = QVBoxLayout(list_panel)
        
        # Create profile list
        self.profile_list = QListWidget()
        self.profile_list.setMinimumWidth(200)
        self.profile_list.currentItemChanged.connect(self.on_profile_selected)
        list_layout.addWidget(QLabel("Available Profiles:"))
        list_layout.addWidget(self.profile_list)
        
        # Add buttons for profile management
        button_layout = QHBoxLayout()
        
        self.new_btn = QPushButton("New")
        self.new_btn.clicked.connect(self.on_new_profile)
        
        self.delete_btn = QPushButton("Delete")
        self.delete_btn.clicked.connect(self.on_delete_profile)
        self.delete_btn.setEnabled(False)  # Disable until a profile is selected
        
        self.default_btn = QPushButton("Set as Default")
        self.default_btn.clicked.connect(self.on_set_default_profile)
        self.default_btn.setEnabled(False)  # Disable until a profile is selected
        
        button_layout.addWidget(self.new_btn)
        button_layout.addWidget(self.delete_btn)
        button_layout.addWidget(self.default_btn)
        
        list_layout.addLayout(button_layout)
        
        # Add the list panel to the main layout
        main_layout.addWidget(list_panel, 1)
        
        # Create profile details panel (right side)
        self.details_panel = QWidget()
        details_layout = QVBoxLayout(self.details_panel)
        
        # Create tabs for different profile settings
        self.tabs = QTabWidget()
        
        # Tab 1: General Settings
        self.general_tab = QWidget()
        general_layout = QFormLayout(self.general_tab)
        
        self.profile_name = QLineEdit()
        general_layout.addRow("Profile Name:", self.profile_name)
        
        self.output_folder = QLineEdit()
        output_browse_btn = QPushButton("Browse...")
        output_browse_btn.clicked.connect(self.on_browse_output_folder)
        
        output_layout = QHBoxLayout()
        output_layout.addWidget(self.output_folder)
        output_layout.addWidget(output_browse_btn)
        
        general_layout.addRow("Default Output Folder:", output_layout)
        
        self.auto_process = QCheckBox("Automatically process matching files")
        general_layout.addRow("", self.auto_process)
        
        self.tabs.addTab(self.general_tab, "General")
        
        # Tab 2: Column Patterns
        self.patterns_tab = QWidget()
        patterns_layout = QVBoxLayout(self.patterns_tab)
        
        patterns_layout.addWidget(QLabel("Column selection patterns:"))
        
        # Table for patterns
        self.patterns_table = QTableWidget(0, 2)
        self.patterns_table.setHorizontalHeaderLabels(["Sheet Pattern", "Columns"])
        self.patterns_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Interactive)
        self.patterns_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        
        patterns_layout.addWidget(self.patterns_table)
        
        # Add/remove pattern buttons
        patterns_btn_layout = QHBoxLayout()
        
        add_pattern_btn = QPushButton("Add Pattern")
        add_pattern_btn.clicked.connect(self.on_add_pattern)
        
        patterns_btn_layout.addWidget(add_pattern_btn)
        patterns_btn_layout.addStretch()
        
        patterns_layout.addLayout(patterns_btn_layout)
        
        self.tabs.addTab(self.patterns_tab, "Column Patterns")
        
        # Tab 3: Watch Folders
        self.watch_tab = QWidget()
        watch_layout = QVBoxLayout(self.watch_tab)
        
        watch_layout.addWidget(QLabel("Watch folders for automatic processing:"))
        
        # List for watch folders
        self.watch_list = QListWidget()
        watch_layout.addWidget(self.watch_list)
        
        # Add/remove folder buttons
        watch_btn_layout = QHBoxLayout()
        
        add_folder_btn = QPushButton("Add Folder")
        add_folder_btn.clicked.connect(self.on_browse_watch_folder)
        
        remove_folder_btn = QPushButton("Remove Selected")
        remove_folder_btn.clicked.connect(self.on_remove_watch_folder)
        
        watch_btn_layout.addWidget(add_folder_btn)
        watch_btn_layout.addWidget(remove_folder_btn)
        watch_btn_layout.addStretch()
        
        watch_layout.addLayout(watch_btn_layout)
        
        self.tabs.addTab(self.watch_tab, "Watch Folders")
        
        # Add the tabs to the details layout
        details_layout.addWidget(self.tabs)
        
        # Add save button
        save_layout = QHBoxLayout()
        
        self.save_btn = QPushButton("Save")
        self.save_btn.clicked.connect(self.on_save_profile)
        self.save_btn.setEnabled(False)  # Disable until a profile is selected or created
        
        self.apply_btn = QPushButton("Apply to Current Data")
        self.apply_btn.clicked.connect(self.on_apply_profile)
        self.apply_btn.setEnabled(False)  # Disable until a profile is selected
        
        # Only enable apply button if we have current data
        if not self.file_data or not self.current_selections:
            self.apply_btn.setEnabled(False)
            self.apply_btn.setToolTip("No data available to apply profile to")
        
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        
        save_layout.addWidget(self.apply_btn)
        save_layout.addStretch()
        save_layout.addWidget(cancel_btn)
        save_layout.addWidget(self.save_btn)
        
        details_layout.addLayout(save_layout)
        
        # Add the details panel to the main layout
        main_layout.addWidget(self.details_panel, 2)
        
        # Disable the details panel until a profile is selected
        self.details_panel.setEnabled(False)
        
        # Load existing profiles
        self.load_profiles()
        
    def load_profiles(self):
        """Load and display available profiles"""
        # Clear the profile list
        self.profile_list.clear()
        
        # Check if we have a profile manager
        if not self.profile_manager:
            return
            
        # Get profiles from the manager
        profiles = self.profile_manager.get_all_profiles()
        default_name = self.profile_manager.default_profile_name
        
        # Add profiles to the list
        for name, profile in sorted(profiles.items()):
            # Create list item
            item = ProfileListItem(profile, name == default_name)
            self.profile_list.addItem(item)
        
    def on_profile_selected(self, current, previous):
        """Handle profile selection"""
        # Enable/disable buttons
        if current:
            self.delete_btn.setEnabled(True)
            self.default_btn.setEnabled(True)
            self.save_btn.setEnabled(True)
            self.apply_btn.setEnabled(bool(self.file_data))
            self.details_panel.setEnabled(True)
            
            # Load profile data
            profile = current.profile
            
            # Update general settings
            self.profile_name.setText(profile.name)
            self.output_folder.setText(profile.output_folder)
            self.auto_process.setChecked(profile.auto_process)
            
            # Load column patterns
            self.load_patterns(profile)
            
            # Load watch folders
            self.watch_list.clear()
            for folder in profile.watch_folders:
                self.watch_list.addItem(folder)
                
            # Store reference to selected profile
            self.selected_profile = profile
            
        else:
            self.delete_btn.setEnabled(False)
            self.default_btn.setEnabled(False)
            self.save_btn.setEnabled(False)
            self.apply_btn.setEnabled(False)
            self.details_panel.setEnabled(False)
            self.selected_profile = None
            
    def load_patterns(self, profile):
        """Load column patterns into the UI"""
        # Clear current patterns
        self.patterns_table.setRowCount(0)
        
        # Add each pattern
        for i, (pattern, columns) in enumerate(profile.column_patterns):
            # Add a new row
            self.patterns_table.insertRow(i)
            
            # Add pattern item
            pattern_item = QTableWidgetItem(pattern)
            self.patterns_table.setItem(i, 0, pattern_item)
            
            # Add columns item (joining all columns with commas)
            columns_text = ", ".join(columns)
            columns_item = QTableWidgetItem(columns_text)
            self.patterns_table.setItem(i, 1, columns_item)
            
            # Add delete button in a delegate
            delete_btn = QPushButton("Delete")
            delete_btn.row = i  # Store row for identifying which to delete
            delete_btn.clicked.connect(lambda checked, row=i: self.on_delete_pattern(row))
            
            # Better way in PyQt5 would be to use setCellWidget but we'll rely on connecting
            # by row number for simplicity here
        
    def on_new_profile(self):
        """Create a new profile"""
        # Create a default name
        base_name = "New Profile"
        name = base_name
        counter = 1
        
        # Generate a unique name
        existing_names = set()
        for i in range(self.profile_list.count()):
            item = self.profile_list.item(i)
            existing_names.add(item.profile.name)
            
        while name in existing_names:
            name = f"{base_name} {counter}"
            counter += 1
            
        # Create the profile
        if self.profile_manager:
            profile = self.profile_manager.create_profile(name)
            
            # Add to the list
            item = ProfileListItem(profile)
            self.profile_list.addItem(item)
            
            # Select the new profile
            self.profile_list.setCurrentItem(item)
            
            # Emit the profiles updated signal
            self.profiles_updated.emit()
            
    def on_delete_profile(self):
        """Delete the selected profile"""
        # Get the selected profile
        current_item = self.profile_list.currentItem()
        if not current_item:
            return
            
        # Confirm deletion
        reply = QMessageBox.question(
            self, 
            "Confirm Deletion", 
            f"Are you sure you want to delete the profile '{current_item.profile.name}'?",
            QMessageBox.Yes | QMessageBox.No, 
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
            
        # Delete the profile
        if self.profile_manager:
            self.profile_manager.delete_profile(current_item.profile.name)
            
            # Remove from the list
            self.profile_list.takeItem(self.profile_list.row(current_item))
            
            # Emit the profiles updated signal
            self.profiles_updated.emit()
            
    def on_set_default_profile(self):
        """Set the selected profile as default"""
        # Get the selected profile
        current_item = self.profile_list.currentItem()
        if not current_item:
            return
            
        # Set as default
        if self.profile_manager:
            self.profile_manager.set_default_profile(current_item.profile.name)
            
            # Update the display
            self.load_profiles()
            
            # Find and select the same profile again
            for i in range(self.profile_list.count()):
                item = self.profile_list.item(i)
                if item.profile.name == current_item.profile.name:
                    self.profile_list.setCurrentItem(item)
                    break
                    
            # Emit the profiles updated signal
            self.profiles_updated.emit()
            
    def on_save_profile(self):
        """Save the current profile"""
        # Get the selected profile
        if not self.selected_profile:
            return
            
        # Update profile from UI
        old_name = self.selected_profile.name
        new_name = self.profile_name.text().strip()
        
        # Validate the name
        if not new_name:
            QMessageBox.warning(self, "Invalid Name", "Please enter a valid profile name.")
            return
            
        # Check if the name changed and if the new name already exists
        if new_name != old_name:
            for i in range(self.profile_list.count()):
                item = self.profile_list.item(i)
                if item.profile.name == new_name:
                    QMessageBox.warning(
                        self, 
                        "Name Conflict", 
                        f"A profile with the name '{new_name}' already exists.\nPlease choose a different name."
                    )
                    return
        
        # Update profile properties
        self.selected_profile.name = new_name
        self.selected_profile.output_folder = self.output_folder.text()
        self.selected_profile.auto_process = self.auto_process.isChecked()
        
        # Update column patterns
        self.selected_profile.column_patterns = []
        for row in range(self.patterns_table.rowCount()):
            pattern_item = self.patterns_table.item(row, 0)
            columns_item = self.patterns_table.item(row, 1)
            
            if pattern_item and columns_item:
                pattern = pattern_item.text().strip()
                columns_text = columns_item.text().strip()
                
                # Parse columns (comma separated)
                columns = [col.strip() for col in columns_text.split(",") if col.strip()]
                
                # Add to profile
                if pattern and columns:
                    self.selected_profile.column_patterns.append((pattern, columns))
        
        # Update watch folders
        self.selected_profile.watch_folders = []
        for i in range(self.watch_list.count()):
            folder = self.watch_list.item(i).text()
            if folder:
                self.selected_profile.watch_folders.append(folder)
        
        # Save the profile
        if self.profile_manager:
            # Handle profile rename
            if new_name != old_name:
                # Rename the profile
                self.profile_manager.rename_profile(old_name, new_name)
            else:
                # Just save the profile
                self.profile_manager.save_profile(self.selected_profile)
                
            # Update the display
            self.load_profiles()
            
            # Find and select the same profile again
            for i in range(self.profile_list.count()):
                item = self.profile_list.item(i)
                if item.profile.name == new_name:
                    self.profile_list.setCurrentItem(item)
                    break
                    
            # Emit the profiles updated signal
            self.profiles_updated.emit()
            
            # Show confirmation
            QMessageBox.information(self, "Profile Saved", f"Profile '{new_name}' has been saved.")
            
    def on_browse_output_folder(self):
        """Browse for output folder"""
        folder = QFileDialog.getExistingDirectory(
            self, 
            "Select Output Folder", 
            self.output_folder.text()
        )
        
        if folder:
            self.output_folder.setText(folder)
            
    def on_browse_watch_folder(self):
        """Browse for watch folder"""
        folder = QFileDialog.getExistingDirectory(
            self, 
            "Select Watch Folder", 
            ""
        )
        
        if folder:
            # Check if already in the list
            for i in range(self.watch_list.count()):
                if self.watch_list.item(i).text() == folder:
                    # Already in the list
                    QMessageBox.information(
                        self, 
                        "Folder Already Added", 
                        f"The folder '{folder}' is already in the watch list."
                    )
                    return
            
            # Add to the list
            self.watch_list.addItem(folder)
            
    def on_remove_watch_folder(self):
        """Remove a watch folder"""
        # Get the selected item
        current_item = self.watch_list.currentItem()
        if current_item:
            self.watch_list.takeItem(self.watch_list.row(current_item))
            
    def on_add_pattern(self):
        """Add a column pattern"""
        # Create a dialog for adding a pattern
        dialog = QDialog(self)
        dialog.setWindowTitle("Add Column Pattern")
        dialog.setMinimumWidth(400)
        
        # Create the layout
        layout = QVBoxLayout(dialog)
        
        # Add a form layout for pattern input
        form_layout = QFormLayout()
        
        # Add sheet pattern input
        sheet_pattern = QLineEdit()
        form_layout.addRow("Sheet Pattern:", sheet_pattern)
        
        # Add help text
        form_layout.addRow("", QLabel("Enter a sheet name or 'file:filename|sheet:sheetname' pattern"))
        
        # Add columns input
        columns_input = QLineEdit()
        form_layout.addRow("Columns:", columns_input)
        
        # Add help text
        form_layout.addRow("", QLabel("Enter column names separated by commas"))
        
        # Add the form to the layout
        layout.addLayout(form_layout)
        
        # Add current data if available
        if self.file_data and self.current_selections:
            # Create a group box for current data
            data_group = QGroupBox("Create From Current Data")
            data_layout = QFormLayout(data_group)
            
            # Create dropdowns for file and sheet
            file_combo = QComboBox()
            sheet_combo = QComboBox()
            
            # Populate file dropdown
            for file_name in sorted(self.file_data.keys()):
                file_combo.addItem(file_name)
                
            # Connect file dropdown to update sheet dropdown
            def update_sheets():
                sheet_combo.clear()
                file_name = file_combo.currentText()
                if file_name in self.file_data:
                    for sheet_name in sorted(self.file_data[file_name].keys()):
                        sheet_combo.addItem(sheet_name)
                        
            # Connect the signal
            file_combo.currentIndexChanged.connect(update_sheets)
            
            # Add to layout
            data_layout.addRow("File:", file_combo)
            data_layout.addRow("Sheet:", sheet_combo)
            
            # Add a button to use current selection
            use_btn = QPushButton("Use Current Selection")
            
            def use_current_selection():
                file_name = file_combo.currentText()
                sheet_name = sheet_combo.currentText()
                
                if (
                    file_name in self.current_selections and 
                    sheet_name in self.current_selections[file_name]
                ):
                    # Get the columns
                    columns = self.current_selections[file_name][sheet_name]
                    
                    # Set the pattern
                    sheet_pattern.setText(f"file:{file_name}|sheet:{sheet_name}")
                    
                    # Set the columns
                    columns_input.setText(", ".join(str(col) for col in columns))
                    
            # Connect the button
            use_btn.clicked.connect(use_current_selection)
            
            # Add the button to the layout
            data_layout.addRow("", use_btn)
            
            # Initialize the sheet dropdown
            update_sheets()
            
            # Add the group to the layout
            layout.addWidget(data_group)
            
        # Add buttons
        button_layout = QHBoxLayout()
        
        ok_btn = QPushButton("OK")
        ok_btn.clicked.connect(dialog.accept)
        
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(dialog.reject)
        
        button_layout.addStretch()
        button_layout.addWidget(cancel_btn)
        button_layout.addWidget(ok_btn)
        
        layout.addLayout(button_layout)
        
        # Show the dialog
        if dialog.exec_() == QDialog.Accepted:
            # Get the pattern and columns
            pattern = sheet_pattern.text().strip()
            columns_text = columns_input.text().strip()
            
            # Validate
            if not pattern:
                QMessageBox.warning(self, "Invalid Pattern", "Please enter a valid sheet pattern.")
                return
                
            if not columns_text:
                QMessageBox.warning(self, "Invalid Columns", "Please enter at least one column.")
                return
                
            # Parse columns (comma separated)
            columns = [col.strip() for col in columns_text.split(",") if col.strip()]
            
            # Add to table
            row = self.patterns_table.rowCount()
            self.patterns_table.insertRow(row)
            
            # Add pattern item
            pattern_item = QTableWidgetItem(pattern)
            self.patterns_table.setItem(row, 0, pattern_item)
            
            # Add columns item
            columns_item = QTableWidgetItem(columns_text)
            self.patterns_table.setItem(row, 1, columns_item)
            
    def on_delete_pattern(self, row):
        """Delete a column pattern"""
        if 0 <= row < self.patterns_table.rowCount():
            self.patterns_table.removeRow(row)
            
    def on_apply_profile(self):
        """Apply the selected profile to current data"""
        # Save current profile first
        self.on_save_profile()
        
        # Set the dialog result
        self.accept()