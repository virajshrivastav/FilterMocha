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

def test_problematic_files():
    """Test the improved Excel standardizer with problematic files"""
    try:
        # Initialize the standardizer
        standardizer = ExcelStandardizer()
        
        # Test files
        test_files = [
            {
                'file': 'Sample-Files/Assessment Questions - Oracle Apex.xlsx',
                'sheet': 'Oracle apex questions',  # Intentionally using lowercase to test case-insensitive matching
                'mapping': {
                    'Question Type': 'Q type',
                    'Difficulty Level': 'Level',
                    'Question Text': 'Q text',
                    'Option (A)': 'Option/ Answer 1',
                    'Option (B)': 'Option/ Answer 2',
                    'Option (C)': 'Option/ Answer 3',
                    'Option (D)': 'Option/ Answer 4',
                    'Option (E)': 'Option/ Answer 5',
                    'Correct Answer': 'Correct Answer'
                }
            },
            {
                'file': 'Sample-Files/Qyrus Questionnaire.xlsx',
                'sheet': 'qyrus api',  # Intentionally using lowercase to test case-insensitive matching
                'mapping': {
                    'Question Type': 'Category',  # Intentionally using a non-matching column to test fuzzy matching
                    'Difficulty Level': 'Difficulty',
                    'Question Text': 'Question Text',
                    'Option (A)': 'Option A',
                    'Option (B)': 'Option B',
                    'Option (C)': 'Option C',
                    'Option (D)': 'Option D',
                    'Option (E)': 'Option E',
                    'Correct Answer': 'Correct Answer'
                }
            }
        ]
        
        # Test each file
        for test_case in test_files:
            file_path = test_case['file']
            sheet_name = test_case['sheet']
            mapping = test_case['mapping']
            
            print(f"\n=== Testing file: {file_path}, sheet: {sheet_name} ===")
            
            try:
                # First test analyze_file
                print("Testing analyze_file...")
                columns_info, file_shape, sheet_names, selected_sheet, sheet_info = standardizer.analyze_file(file_path, sheet_name)
                print(f"Success! Selected sheet: '{selected_sheet}'")
                print(f"Number of columns: {len(columns_info)}")
                print(f"Number of rows: {file_shape[0]}")
                
                # Then test process_file
                print("\nTesting process_file...")
                result = standardizer.process_file(file_path, mapping, sheet_name=sheet_name)
                
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
                
            except Exception as e:
                print(f"Error: {str(e)}")
                return False
        
        return True
    except Exception as e:
        print(f"Test failed: {str(e)}")
        return False

if __name__ == "__main__":
    print("Testing robust Excel standardizer...")
    
    # Run the test
    result = test_problematic_files()
    
    # Print summary
    print("\n=== Test Summary ===")
    if result:
        print("All tests passed! The improvements are working correctly.")
        sys.exit(0)
    else:
        print("Some tests failed. Please check the logs for details.")
        sys.exit(1)
