# Excel Data Extractor

A MacOS application for extracting and merging selected data from multiple Excel files contained in a zip archive.

## Overview

This repository contains two implementations of the Excel Data Extractor:

1. **PyQt5 Desktop Application** (`excel_extractor_qt.py`) - A native MacOS desktop application with a graphical user interface
2. **Command-Line Version** (`excel_extractor_cli.py`) - A simple command-line interface for the same functionality

The application allows you to:
- Extract Excel files from a ZIP archive
- View the data in each Excel file by sheet
- Select specific columns from each sheet using checkboxes (GUI) or interactive selection (CLI)
- Merge the selected data into a new Excel file
- Generate a summary of extracted data

## Requirements

### For the PyQt5 GUI Application

- Python 3.6 or higher
- PyQt5
- pandas
- xlwt
- openpyxl

### For the Command-Line Version

- Python 3.6 or higher
- pandas
- xlwt
- openpyxl

## Installation

1. Clone or download this repository
2. Install the required dependencies:

```bash
pip install pyqt5 pandas xlwt openpyxl
```

## Usage

### PyQt5 GUI Application

Run the GUI application with:

```bash
python excel_extractor_qt.py
```

The application will open a window with the following workflow:

1. **Upload ZIP**: Browse for a ZIP file containing Excel files
2. **Select Data**: Choose which columns to extract from each Excel file
3. **Generate Output**: Specify the output filename and location, then generate the merged Excel file

### Command-Line Version

Run the command-line version with:

```bash
python excel_extractor_cli.py [zip_file] [output_file]
```

For example:

```bash
python excel_extractor_cli.py data.zip merged_output.xls
```

The CLI version will:

1. Extract Excel files from the ZIP archive
2. Display available columns for each file and sheet
3. Prompt you to select columns
4. Process and generate the merged Excel file

## MacOS Optimizations

This application is optimized for MacOS with:

- Native file dialogs in the GUI version
- Proper handling of MacOS file paths
- Compatible with MacOS system themes and styles
- Appropriate handling of temporary directories

## Troubleshooting

If you encounter issues:

1. Make sure all dependencies are installed
2. Verify your ZIP file contains valid Excel (.xlsx or .xls) files
3. For password-protected Excel files, remove the password protection before processing
4. For large Excel files, ensure your system has sufficient memory