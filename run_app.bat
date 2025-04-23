@echo off
color E7
echo Excel File Standardizer for FilterMocha
echo =====================================

echo Starting Excel File Standardizer...
echo.
echo Features:
echo - Modern web-based UI with orange theme
echo - Drag and drop file upload
echo - Interactive column mapping
echo - File splitting option (enabled after file upload)
echo - Download-only functionality for output files
echo - Download all files as ZIP option
echo - Complete standard format with all fields
echo - Sheet selection for multi-sheet Excel files
echo - Topics field not mandatory
echo - iMocha branding
echo - Question numbers in warning files
echo - Optimized Excel output with increased column widths
echo - Support for FIB and DESC question types
echo - Standardized file naming conventions
echo.
echo Once the server is running, open your browser and go to:
echo http://localhost:5051
echo.
echo Press Ctrl+C to stop the server when you're done.
echo.

REM Check if Flask is installed
python -c "import flask" 2>nul
if %errorlevel% neq 0 (
    echo Installing required packages...
    pip install flask flask-cors pandas openpyxl
)

REM Check if xlsxwriter is installed
python -c "import xlsxwriter" 2>nul
if %errorlevel% neq 0 (
    echo Installing xlsxwriter package...
    pip install xlsxwriter
)

REM Create required directories
if not exist "uploads" mkdir uploads
if not exist "Processed-Files" mkdir Processed-Files

REM Start the web application with optimized settings
python fixed_app.py
