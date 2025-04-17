"""
Excel Data Extractor - Profile Manager
This module handles the creation, loading, saving, and application of extraction profiles
for the Excel Data Extractor application.

Profiles allow users to save column selection patterns for repeated extraction tasks,
and can be associated with watched folders for automated processing.
"""

import os
import json
import re
from pathlib import Path
from typing import Dict, List, Optional, Any, Set


def get_app_data_dir() -> str:
    """Get the application data directory for storing profiles and settings"""
    # Use platform-specific app data locations
    if os.name == 'nt':  # Windows
        app_data = os.environ.get('APPDATA', '')
        if not app_data:
            app_data = os.path.expanduser("~\\AppData\\Roaming")
        base_dir = os.path.join(app_data, "Excel Data Extractor")
    elif os.name == 'posix':  # macOS and Linux
        if os.path.exists('/Applications'):  # macOS
            base_dir = os.path.expanduser("~/Library/Application Support/Excel Data Extractor")
        else:  # Linux
            base_dir = os.path.expanduser("~/.config/excel-data-extractor")
    else:
        # Fallback option
        base_dir = os.path.expanduser("~/.excel-data-extractor")
    
    # Create directory if it doesn't exist
    if not os.path.exists(base_dir):
        os.makedirs(base_dir)
        
    # Create profiles directory
    profiles_dir = os.path.join(base_dir, "profiles")
    if not os.path.exists(profiles_dir):
        os.makedirs(profiles_dir)
        
    return base_dir


class ExtractionProfile:
    """Represents a saved extraction profile with column selection patterns and settings"""
    
    def __init__(self, name: str = "New Profile"):
        """Initialize a new extraction profile"""
        self.name = name
        self.column_patterns = []  # List of (sheet_pattern, column_list) tuples
        self.watch_folders = []    # List of folder paths to watch for new files
        self.output_folder = ""    # Default output folder for this profile
        self.auto_process = False  # Whether to automatically process matching files
        
    def to_dict(self) -> Dict[str, Any]:
        """Convert profile to dictionary for JSON serialization"""
        return {
            "name": self.name,
            "column_patterns": self.column_patterns,
            "watch_folders": self.watch_folders,
            "output_folder": self.output_folder,
            "auto_process": self.auto_process
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'ExtractionProfile':
        """Create profile from dictionary (from JSON)"""
        profile = cls(data.get("name", "Unnamed Profile"))
        profile.column_patterns = data.get("column_patterns", [])
        profile.watch_folders = data.get("watch_folders", [])
        profile.output_folder = data.get("output_folder", "")
        profile.auto_process = data.get("auto_process", False)
        return profile
    
    def add_column_pattern(self, sheet_pattern: str, columns: List[str]) -> None:
        """Add a column selection pattern"""
        # Don't add duplicate patterns
        for pattern, cols in self.column_patterns:
            if pattern == sheet_pattern:
                # Update existing pattern with these columns
                # Get unique columns
                unique_cols = list(set(cols + columns))
                # Replace the existing pattern
                self.column_patterns.remove((pattern, cols))
                self.column_patterns.append((pattern, unique_cols))
                return
                
        # Add as a new pattern
        self.column_patterns.append((sheet_pattern, columns))
    
    def add_file_selection(self, file_name: str, sheet_name: str, columns: List[str]) -> None:
        """Add a specific file selection"""
        # For file-specific patterns, we use "file:filename|sheet:sheetname" format
        pattern = f"file:{file_name}|sheet:{sheet_name}"
        self.add_column_pattern(pattern, columns)
        
    def add_watch_folder(self, folder_path: str) -> None:
        """Add a folder to watch for new files"""
        if folder_path and folder_path not in self.watch_folders:
            self.watch_folders.append(folder_path)
            
    def remove_watch_folder(self, folder_path: str) -> None:
        """Remove a watched folder"""
        if folder_path in self.watch_folders:
            self.watch_folders.remove(folder_path)
    
    def _pattern_matches_sheet(self, pattern: str, file_name: str, sheet_name: str) -> bool:
        """Check if a pattern matches a file and sheet name"""
        # Handle specific file and sheet patterns
        if pattern.startswith("file:"):
            # Pattern format: "file:filename|sheet:sheetname"
            parts = pattern.split("|")
            if len(parts) != 2:
                return False
                
            file_pattern = parts[0].replace("file:", "").strip()
            sheet_pattern = parts[1].replace("sheet:", "").strip()
            
            # Simple exact matching for now
            return file_name == file_pattern and sheet_name == sheet_pattern
        
        # For general patterns, just check if the sheet name matches
        # (In the future, this could support regex or glob patterns)
        return sheet_name == pattern
    
    def match_to_new_files(self, file_data: Dict[str, Dict[str, Any]]) -> Dict[str, Dict[str, List[str]]]:
        """
        Apply this profile's selection patterns to new file data.
        
        Args:
            file_data: The new file data structure {file_name: {sheet_name: dataframe}}
            
        Returns:
            A selection structure {file_name: {sheet_name: [columns]}}
        """
        selections = {}
        
        # Apply each pattern to the file data
        for pattern, columns in self.column_patterns:
            for file_name, sheets in file_data.items():
                for sheet_name, df in sheets.items():
                    if self._pattern_matches_sheet(pattern, file_name, sheet_name):
                        # Initialize the structure if needed
                        if file_name not in selections:
                            selections[file_name] = {}
                            
                        if sheet_name not in selections[file_name]:
                            selections[file_name][sheet_name] = []
                            
                        # Add the columns if they exist in this dataframe
                        for col in columns:
                            if col in df.columns and col not in selections[file_name][sheet_name]:
                                selections[file_name][sheet_name].append(col)
                            
        return selections


class ProfileManager:
    """Manager for handling multiple extraction profiles"""
    
    def __init__(self):
        """Initialize the profile manager"""
        self.app_data_dir = get_app_data_dir()
        self.profiles_dir = os.path.join(self.app_data_dir, "profiles")
        self.profiles = {}  # Dictionary of name -> profile objects
        self.default_profile_name = ""
        
        # Make sure the directories exist
        if not os.path.exists(self.profiles_dir):
            os.makedirs(self.profiles_dir)
        
        # Load existing profiles if any
        self.load_all_profiles()
        self.load_settings()
    
    def create_profile(self, name: str) -> ExtractionProfile:
        """Create a new profile"""
        # Generate unique name if needed
        base_name = name
        counter = 1
        while name in self.profiles:
            name = f"{base_name} ({counter})"
            counter += 1
            
        # Create and store the profile
        profile = ExtractionProfile(name)
        self.profiles[name] = profile
        
        # Save the profile to disk
        self.save_profile(profile)
        
        return profile
    
    def save_profile(self, profile: ExtractionProfile) -> bool:
        """Save a profile to disk"""
        try:
            # Ensure the profile has a name
            if not profile.name:
                profile.name = "Unnamed Profile"
                
            # Make the name safe for filenames
            safe_name = re.sub(r'[^\w\-_\. ]', '_', profile.name)
            
            # Create the profile file path
            file_path = os.path.join(self.profiles_dir, f"{safe_name}.json")
            
            # Save the profile as JSON
            with open(file_path, 'w') as f:
                json.dump(profile.to_dict(), f, indent=2)
                
            # Update our profiles dictionary
            self.profiles[profile.name] = profile
            
            return True
            
        except Exception as e:
            print(f"Error saving profile: {str(e)}")
            return False
    
    def load_profile(self, file_path: str) -> Optional[ExtractionProfile]:
        """Load a profile from a file"""
        try:
            with open(file_path, 'r') as f:
                data = json.load(f)
                
            # Create profile from the data
            profile = ExtractionProfile.from_dict(data)
            
            # Store in the profiles dictionary
            self.profiles[profile.name] = profile
            
            return profile
            
        except Exception as e:
            print(f"Error loading profile from {file_path}: {str(e)}")
            return None
    
    def load_all_profiles(self) -> None:
        """Load all profiles from the profiles directory"""
        # Clear existing profiles
        self.profiles = {}
        
        # Look for profile files
        for file_name in os.listdir(self.profiles_dir):
            if file_name.endswith(".json"):
                file_path = os.path.join(self.profiles_dir, file_name)
                profile = self.load_profile(file_path)
                if profile:
                    print(f"Loaded profile: {profile.name}")
    
    def delete_profile(self, name: str) -> bool:
        """Delete a profile"""
        if name not in self.profiles:
            return False
            
        # Get the profile
        profile = self.profiles[name]
        
        # Create the profile file path
        safe_name = re.sub(r'[^\w\-_\. ]', '_', profile.name)
        file_path = os.path.join(self.profiles_dir, f"{safe_name}.json")
        
        # Remove the file if it exists
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                
            # Remove from our profiles dictionary
            del self.profiles[name]
            
            # If this was the default profile, clear that setting
            if self.default_profile_name == name:
                self.default_profile_name = ""
                self.save_settings()
                
            return True
            
        except Exception as e:
            print(f"Error deleting profile: {str(e)}")
            return False
    
    def rename_profile(self, old_name: str, new_name: str) -> bool:
        """Rename a profile"""
        if old_name not in self.profiles or new_name in self.profiles:
            return False
            
        # Get the profile
        profile = self.profiles[old_name]
        
        # Create the old profile file path
        safe_old_name = re.sub(r'[^\w\-_\. ]', '_', old_name)
        old_file_path = os.path.join(self.profiles_dir, f"{safe_old_name}.json")
        
        # Try to remove the old file
        try:
            if os.path.exists(old_file_path):
                os.remove(old_file_path)
                
            # Update the profile name
            profile.name = new_name
            
            # Save with new name
            self.save_profile(profile)
            
            # Update our profiles dictionary
            del self.profiles[old_name]
            self.profiles[new_name] = profile
            
            # Update default profile if needed
            if self.default_profile_name == old_name:
                self.default_profile_name = new_name
                self.save_settings()
                
            return True
            
        except Exception as e:
            print(f"Error renaming profile: {str(e)}")
            return False
    
    def set_default_profile(self, name: str) -> bool:
        """Set the default profile"""
        if name not in self.profiles and name != "":
            return False
            
        # Update the default profile name
        self.default_profile_name = name
        
        # Save the settings
        return self.save_settings()
    
    def get_default_profile(self) -> Optional[ExtractionProfile]:
        """Get the default profile"""
        if not self.default_profile_name or self.default_profile_name not in self.profiles:
            return None
            
        return self.profiles[self.default_profile_name]
    
    def get_all_profiles(self) -> Dict[str, ExtractionProfile]:
        """Get all profiles"""
        return self.profiles
    
    def get_profile(self, name: str) -> Optional[ExtractionProfile]:
        """Get a profile by name"""
        if name not in self.profiles:
            return None
            
        return self.profiles[name]
    
    def save_settings(self) -> bool:
        """Save settings like default profile name"""
        try:
            # Create the settings file path
            settings_path = os.path.join(self.app_data_dir, "settings.json")
            
            # Create the settings dictionary
            settings = {
                "default_profile": self.default_profile_name
            }
            
            # Save the settings as JSON
            with open(settings_path, 'w') as f:
                json.dump(settings, f, indent=2)
                
            return True
            
        except Exception as e:
            print(f"Error saving settings: {str(e)}")
            return False
    
    def load_settings(self) -> bool:
        """Load settings"""
        try:
            # Create the settings file path
            settings_path = os.path.join(self.app_data_dir, "settings.json")
            
            # Check if the settings file exists
            if not os.path.exists(settings_path):
                return False
                
            # Load the settings from JSON
            with open(settings_path, 'r') as f:
                settings = json.load(f)
                
            # Update settings
            self.default_profile_name = settings.get("default_profile", "")
            
            return True
            
        except Exception as e:
            print(f"Error loading settings: {str(e)}")
            return False