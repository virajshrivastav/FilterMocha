# Excel File Standardizer for FilterMocha (Official Backup)

This is the official backup of the Excel File Standardizer application for FilterMocha. The application allows users to upload Excel files, map columns to a standard format, and process the files according to specific requirements.

## Features

- Modern web-based UI with orange theme
- Drag and drop file upload
- Interactive column mapping
- File splitting option (enabled only after file upload)
- Download-only functionality for output files
- Complete standard format with all fields
- Sheet selection for multi-sheet Excel files
- Topics field not mandatory
- iMocha branding
- Question numbers in warning files for easy identification
- Optimized Excel output with increased column widths and better formatting
- Support for additional question types: Fill in the Blank (FIB) and Descriptive (DESC)
- Standardized file naming conventions for output files
- Download all files as ZIP option for batch downloading

## Files

- `fixed_app.py`: The main Flask application
- `excel_standardizer_improved.py`: The core Excel processing logic
- `run_fixed_app.bat`: Batch file to start the application
- `templates/simple_upload.html`: The main UI template

## Requirements

- Python 3.6+
- Flask
- Pandas
- OpenPyXL
- XlsxWriter

## Directories

- `uploads`: Temporary storage for uploaded files
- `Processed-Files`: Output directory for processed files
- `templates`: HTML templates for the web UI
