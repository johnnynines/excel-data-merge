# Excel Data Extractor Tutorial

This tutorial will walk you through using the Excel Data Extractor application on macOS to extract and merge data from multiple Excel files.

## Prerequisites

Before you begin, make sure you have:
- macOS 10.12 or higher
- Python 3.6 or higher
- Required Python packages installed (see README.md for installation instructions)

## Creating a Test ZIP File

First, let's create a sample ZIP file with Excel files to test the application:

1. Create a few sample Excel files (you can use your existing files or create new ones)
2. Each Excel file should have at least one sheet with some data
3. Select all the Excel files in Finder
4. Right-click and choose "Compress [number] items" to create a ZIP archive
5. macOS will create a file named "Archive.zip" by default

## Using the wxPython GUI Version

### Step 1: Launch the Application
```bash
python excel_extractor_wx.py
```

### Step 2: Select the ZIP File
1. In the "Upload ZIP" tab, click "Browse..."
2. Navigate to and select your ZIP file (e.g., "Archive.zip")
3. Click "Open"
4. Click the "Process ZIP File" button
5. Wait for the extraction and analysis to complete

### Step 3: Select Columns to Extract
1. The application will automatically switch to the "Select Data" tab
2. Navigate through the tabs for each Excel file and sheet
3. For each sheet, use the checkboxes to select which columns you want to include
   - Use "Select All" to select all columns in a sheet
   - Use "Deselect All" to clear all selections for a sheet
4. Click "Continue to Output" when you've finished making your selections

### Step 4: Generate the Output File
1. In the "Generate Output" tab, enter a name for your output file (e.g., "merged_data")
2. Click "Browse..." next to the save location field
3. Select a folder where you want to save the output file
4. Click "Process and Generate Excel File"
5. Wait for the processing to complete

### Step 5: Review the Results
1. A message will appear confirming the file has been saved
2. Navigate to the location you selected to find your merged Excel file
3. Open the file to verify the data has been extracted correctly
4. Note the "Summary" sheet which provides details on what data was extracted

## Using the Command-Line Version

### Basic Usage
```bash
python excel_extractor_cli.py path_to_zip_file.zip output_filename.xls
```

### Example with Interactive Selection
```bash
python excel_extractor_cli.py Archive.zip merged_output.xls
```

Follow the interactive prompts:
1. For each sheet, you'll see column numbers and names
2. Enter column numbers separated by commas (e.g., "1,3,5") to select specific columns
3. Type "all" to select all columns in a sheet
4. Type "done" when you've finished selecting columns for a sheet

### Command-Line Help
To see all available options:
```bash
python excel_extractor_cli.py --help
```

## Tips for Effective Use

1. **Organizing Input Files**: Name your Excel files descriptively, as these names will appear in the application and in the output file
   
2. **Column Selection Strategy**: 
   - Be selective about which columns you extract to keep the output manageable
   - If columns have similar data across sheets, select them consistently for better comparison

3. **Output File Organization**:
   - The output Excel file will have one sheet per input sheet that had columns selected
   - Sheet names in the output follow the pattern "filename_sheetname"
   - A Summary sheet is included that lists all extracted columns by file and sheet

4. **Processing Large Files**:
   - For very large Excel files, use the command-line version which may be more efficient
   - Process files in smaller batches if you encounter memory issues

## Troubleshooting

If you encounter any issues, refer to the "Troubleshooting" section in the README.md file.