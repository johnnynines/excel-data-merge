#!/usr/bin/env python3
"""
Excel Data Extractor - macOS Application Launcher
This script launches the PyQt5 macOS application for extracting and merging
data from multiple Excel files contained in a ZIP archive.
"""

import sys
import os
import platform

# Check if running on macOS
is_macos = platform.system() == "Darwin"

def main():
    """
    Main entry point for the application:
    1. On macOS: Launch the PyQt5 application
    2. On other platforms: Launch the PyQt5 application with a warning
    """
    try:
        # Import the PyQt5 application
        from excel_extractor_qt import main as launch_qt_app
        
        # Display platform information
        if not is_macos:
            print("WARNING: This application is optimized for macOS.")
            print("Some features may not work correctly on other operating systems.")
        
        # Launch the PyQt5 application
        launch_qt_app()
        
    except ImportError as e:
        print(f"ERROR: Could not import required modules: {e}")
        print("Please make sure PyQt5, pandas, and xlwt are installed.")
        sys.exit(1)
    except Exception as e:
        print(f"ERROR: An unexpected error occurred: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
