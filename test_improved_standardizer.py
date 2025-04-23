import pandas as pd
import sys
import os
import logging
from excel_standardizer_improved import ExcelStandardizer

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler()
    ]
)
logger = logging.getLogger()

def test_sheet_name_matching():
    """Test the improved sheet name matching"""
    try:
        # Initialize the standardizer
        standardizer = ExcelStandardizer()
        
        # Load the Excel file
        excel_file = 'Sample-Files/Assessment Questions - Oracle Apex.xlsx'
        
        # Test with different sheet name variations
        test_cases = [
            'Oracle Apex Questions',  # Exact match
            'oracle apex questions',  # Case-insensitive match
            'Oracle apex Questions',  # Mixed case
            'Oracle apex',            # Partial match
            'apex questions',         # Partial match
            'Oracle Questions'        # Fuzzy match
        ]
        
        print("\n=== Testing Sheet Name Matching ===")
        for i, sheet_name in enumerate(test_cases):
            print(f"\nTest {i+1}: '{sheet_name}'")
            
            try:
                # Try to analyze the file with the sheet name
                columns_info, file_shape, sheet_names, selected_sheet, sheet_info = standardizer.analyze_file(excel_file, sheet_name)
                print(f"Success! Selected sheet: '{selected_sheet}'")
                print(f"Number of columns: {len(columns_info)}")
                print(f"Number of rows: {file_shape[0]}")
            except Exception as e:
                print(f"Error: {str(e)}")
        
        return True
    except Exception as e:
        print(f"Test failed: {str(e)}")
        return False

def test_column_mapping():
    """Test the improved column mapping"""
    try:
        # Initialize the standardizer
        standardizer = ExcelStandardizer()
        
        # Load the Excel file
        excel_file = 'Sample-Files/Assessment Questions - Oracle Apex.xlsx'
        
        # Get sheet information
        columns_info, file_shape, sheet_names, selected_sheet, sheet_info = standardizer.analyze_file(excel_file, 'Oracle Apex Questions')
        
        # Create a simple mapping config with variations
        mapping_config = {
            'Question Type': 'Q type ',  # Trailing space
            'Difficulty Level': 'Level',  # Exact match
            'Question Text': 'Q text',    # Different name but similar
            'Option (A)': 'Option/ Answer 1',  # Different format
            'Option (B)': 'Option/ Answer 2',
            'Correct Answer': 'Correct Answer'  # Exact match
        }
        
        print("\n=== Testing Column Mapping ===")
        
        try:
            # Process the file with the mapping
            result = standardizer.process_file(excel_file, mapping_config, sheet_name='Oracle Apex Questions')
            
            print(f"Success! Processed file with {len(result['output_files'])} output files")
            print(f"Output files: {[os.path.basename(f) for f in result['output_files']]}")
            
            if result['errors']:
                print(f"Warnings/Errors: {len(result['errors'])}")
                for i, error in enumerate(result['errors'][:5]):
                    print(f"  {i+1}. {error}")
                if len(result['errors']) > 5:
                    print(f"  ... and {len(result['errors']) - 5} more")
            else:
                print("No errors reported")
                
            return True
        except Exception as e:
            print(f"Error processing file: {str(e)}")
            return False
            
    except Exception as e:
        print(f"Test failed: {str(e)}")
        return False

if __name__ == "__main__":
    print("Testing improved Excel standardizer...")
    
    # Run the tests
    sheet_test_result = test_sheet_name_matching()
    column_test_result = test_column_mapping()
    
    # Print summary
    print("\n=== Test Summary ===")
    print(f"Sheet name matching: {'PASSED' if sheet_test_result else 'FAILED'}")
    print(f"Column mapping: {'PASSED' if column_test_result else 'FAILED'}")
    
    if sheet_test_result and column_test_result:
        print("\nAll tests passed! The improvements are working correctly.")
        sys.exit(0)
    else:
        print("\nSome tests failed. Please check the logs for details.")
        sys.exit(1)
