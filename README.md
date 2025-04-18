# Excel Data Extractor

A native macOS application for extracting and merging selected data from multiple Excel files contained in a ZIP archive.

## Overview

This repository contains two implementations of the Excel Data Extractor:

1. **PyQt5 Desktop Application** (`excel_extractor_qt.py`) - A native macOS desktop application with a graphical user interface, optimized specifically for macOS with proper dark mode support and Retina display compatibility
2. **Command-Line Version** (`excel_extractor_cli.py`) - A command-line interface for batch processing or script integration

The application allows you to:
- Extract Excel files from a ZIP archive
- View the data in each Excel file by sheet
- Select specific columns from each sheet using checkboxes
- Merge the selected data into a new Excel file
- Generate a summary of extracted data

## Requirements

- macOS 10.12 or higher
- Python 3.6 or higher
- PyQt5 (for GUI version)
- pandas
- xlwt
- xlrd
- openpyxl

## Installation

1. Clone or download this repository
2. Install the required dependencies:

```bash
pip install PyQt5 pandas xlwt openpyxl
```

## Detailed Installation & Usage Instructions for macOS

### Step 1: Install Python (if not already installed)
1. Visit [python.org](https://www.python.org/downloads/macos/) and download the latest Python installer for macOS
2. Open the downloaded .pkg file and follow the installation instructions
3. Verify the installation by opening Terminal and typing:
   ```bash
   python3 --version
   ```

### Step 2: Download the Application
1. Download this repository:
   ```bash
   git clone https://github.com/yourusername/excel-data-extractor.git
   cd excel-data-extractor
   ```
   Or download and extract the ZIP file from the repository

### Step 3: Set Up a Virtual Environment (recommended)
1. Create a virtual environment:
   ```bash
   python3 -m venv venv
   ```
2. Activate the virtual environment:
   ```bash
   source venv/bin/activate
   ```

### Step 4: Install Dependencies
Install all required packages:
```bash
pip install PyQt5 pandas xlwt openpyxl
```

### Step 5: Run the Application

#### GUI Version (PyQt5)
Launch the GUI application with:
```bash
python excel_extractor_qt.py
```

#### Alternative Launch Method
You can also use the app.py launcher, which will automatically select the best version for your system:
```bash
python app.py
```

#### Command-Line Version
Run the command-line version with:
```bash
python excel_extractor_cli.py input.zip output.xls
```

For help with command-line options:
```bash
python excel_extractor_cli.py --help
```

### Application Workflow

#### GUI Version Workflow
Once the GUI application is running, follow these steps:

1. **Upload ZIP File**:
   - Click "Browse..." to select a ZIP file containing Excel files
   - Click "Process ZIP File" to extract and analyze the Excel files
   - The extraction process will be displayed in the log area

2. **Select Data to Extract**:
   - Navigate through the tabs for each Excel file and sheet
   - Use checkboxes to select which columns you want to extract
   - Use "Select All" or "Deselect All" buttons for quick selection
   - Click "Continue to Output" when you've made your selections

3. **Generate Output File**:
   - Enter a name for the output Excel file
   - Choose a save location using the directory picker
   - Click "Process and Generate Excel File"
   - The merged Excel file will be created at your chosen location

4. **Review Results**:
   - After processing, the application will show a success message
   - You can choose to process another ZIP file or close the application

#### Command-Line Version Workflow

The command-line version follows this workflow:

1. **File Processing**:
   - The application extracts Excel files from the specified ZIP file
   - Files are analyzed and sheet information is displayed

2. **Interactive Column Selection**:
   - For each sheet in each file, you'll see a data preview and column list
   - Enter column numbers (comma-separated) or 'all' to select all columns
   - You can add more columns or type 'done' to finish selecting

3. **Result Generation**:
   - Selected data is processed and merged
   - An output Excel file is created at the specified path
   - A summary of the process is shown

## macOS Optimizations

This application is optimized for macOS with:

- Native macOS look and feel using PyQt5
- Dark mode support with automatic adaptation to system preferences
- Retina display compatibility with high-DPI scaling
- Native file dialogs for familiar user experience
- Proper handling of macOS file paths
- Compatible with macOS system themes and styles
- Appropriate handling of temporary directories
- MacOS application identity and document associations

## Troubleshooting

If you encounter issues:

1. Make sure all dependencies are installed
2. Verify your ZIP file contains valid Excel (.xlsx or .xls) files
3. For password-protected Excel files, remove the password protection before processing
4. For large Excel files, ensure your system has sufficient memory

### Common Issues and Solutions

#### PyQt5 Installation Issues
If you have trouble installing PyQt5, try:
```bash
pip install --upgrade pip
pip install PyQt5
```

For macOS, you might need to install Qt dependencies using Homebrew:
```bash
brew install qt
```

#### wxPython Installation Problems (Alternative GUI)
If you want to try the alternative wxPython version and have installation issues:
```bash
pip install -U wxPython
```

For macOS, you might need to install some prerequisites:
```bash
brew install pkg-config
```

#### Cannot Open ZIP File
- Make sure the ZIP file is not corrupted
- Verify you have read permissions for the file
- Try creating a new ZIP file with the Excel files

#### Excel Files Not Detected
- Ensure files have proper .xlsx or .xls extensions
- Check if Excel files are in root of ZIP or nested in folders
- Verify Excel files are not empty or corrupted

#### Processing Large Files
- For very large Excel files, increase your system memory
- Process files in smaller batches if needed
- Close other memory-intensive applications