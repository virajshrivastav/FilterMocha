import pandas as pd
import sys
import os
from excel_standardizer_improved import ExcelStandardizer

try:
    # Initialize the standardizer
    standardizer = ExcelStandardizer()
    
    # Load the Excel file
    excel_file = 'Sample-Files/Assessment Questions - Oracle Apex.xlsx'
    
    # Test with lowercase sheet name
    sheet_name = 'oracle apex questions'
    print(f'\nTesting with lowercase sheet name: "{sheet_name}"')
    
    try:
        # Try to analyze the file with the lowercase sheet name
        columns_info, file_shape, sheet_names, selected_sheet, sheet_info = standardizer.analyze_file(excel_file, sheet_name)
        print(f'Success! Sheet was found using case-insensitive matching.')
        print(f'Selected sheet: {selected_sheet}')
        print(f'Number of columns: {len(columns_info)}')
        print(f'Number of rows: {file_shape[0]}')
        
        # Print the first few column names
        print('\nSample columns:')
        for i, col_info in enumerate(columns_info[:5]):
            print(f"{i+1}. {col_info['name']}")
        
    except Exception as e:
        print(f'Error: {str(e)}')
    
except Exception as e:
    print(f'Error: {str(e)}')
    sys.exit(1)
