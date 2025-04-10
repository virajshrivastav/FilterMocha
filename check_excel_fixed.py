import pandas as pd
import sys

try:
    # Load the Excel file
    excel_file = 'Sample-Files/Assessment Questions - Oracle Apex.xlsx'
    xl = pd.ExcelFile(excel_file)
    
    # Print sheet names
    print('Sheet names in the file:')
    for i, sheet in enumerate(xl.sheet_names):
        print(f'{i+1}. {sheet}')
    
    # Try to read the specific sheet
    sheet_name = 'Oracle Apex Questions'  # Corrected case
    print(f'\nAttempting to read sheet: {sheet_name}')
    
    # First check if the sheet exists
    if sheet_name in xl.sheet_names:
        print(f'Sheet "{sheet_name}" exists in the file')
        
        # Try to read a few rows to see if there's an issue
        try:
            df = pd.read_excel(excel_file, sheet_name=sheet_name, nrows=5)
            print(f'Successfully read 5 rows from the sheet')
            print(f'Columns: {df.columns.tolist()}')
            print(f'Shape: {df.shape}')
            
            # Check for any problematic data types or values
            print("\nColumn data types:")
            print(df.dtypes)
            
            # Check for any NaN or problematic values
            print("\nSample data (first 2 rows):")
            print(df.head(2))
            
        except Exception as e:
            print(f'Error reading sheet: {str(e)}')
    else:
        print(f'Sheet "{sheet_name}" does not exist in the file')
        print(f'Available sheets: {xl.sheet_names}')
    
except Exception as e:
    print(f'Error: {str(e)}')
    sys.exit(1)
