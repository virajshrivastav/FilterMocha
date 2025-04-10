import pandas as pd
import sys
import os
import logging
import traceback
from excel_standardizer_improved import ExcelStandardizer

# Configure logging
logging.basicConfig(
    level=logging.DEBUG,  # Set to DEBUG for maximum detail
    format='%(asctime)s - %(levelname)s - %(name)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('test_problematic_file.log')  # Also log to a file
    ]
)
logger = logging.getLogger('test_problematic_file')

def test_file_processing(file_path, sheet_name):
    """Test processing a problematic file with detailed error tracing"""
    try:
        logger.info(f"Testing file: {file_path}, sheet: {sheet_name}")
        
        # Create a standardizer instance
        standardizer = ExcelStandardizer()
        
        # Step 1: Analyze the file
        logger.info("Step 1: Analyzing file...")
        try:
            columns_info, file_shape, sheet_names, selected_sheet, sheet_info = standardizer.analyze_file(file_path, sheet_name)
            logger.info(f"Analysis successful! Selected sheet: {selected_sheet}")
            logger.info(f"Found {len(columns_info)} columns, {file_shape[0]} rows, {len(sheet_names)} sheets")
            logger.info(f"Columns: {[col['name'] for col in columns_info]}")
        except Exception as e:
            logger.error(f"Error analyzing file: {str(e)}")
            logger.error(traceback.format_exc())
            return
        
        # Step 2: Create a mapping configuration
        logger.info("Step 2: Creating mapping configuration...")
        mapping_config = {
            'Question Type': 'Q type',
            'Difficulty Level': 'Level',
            'Question Text': 'Q text',
            'Option (A)': 'Option/ Answer 1',
            'Option (B)': 'Option/ Answer 2',
            'Option (C)': 'Option/ Answer 3',
            'Option (D)': 'Option/ Answer 4',
            'Option (E)': 'Option/ Answer 5',
            'Option (F)': 'Option/ Answer 6',
            'Correct Answer': 'Correct Answer',
            'Topics': 'Topic',
            'Author': "Author's Email id(optional)"
        }
        logger.info(f"Mapping configuration: {mapping_config}")
        
        # Step 3: Process the file
        logger.info("Step 3: Processing file...")
        try:
            # Create a split configuration
            split_config = {
                'enabled': True,
                'column': 'SubSkill'
            }
            
            # Process the file
            result = standardizer.process_file(file_path, mapping_config, split_config, None, sheet_name)
            logger.info(f"Processing successful! Output files: {result['output_files']}")
            
            # Check for errors
            if result['errors']:
                logger.warning(f"Warnings/Errors: {len(result['errors'])}")
                for i, error in enumerate(result['errors']):
                    logger.warning(f"  {i+1}. {error}")
            else:
                logger.info("No errors reported")
        except Exception as e:
            logger.error(f"Error processing file: {str(e)}")
            logger.error(traceback.format_exc())
            return
        
        logger.info("Test completed successfully!")
    
    except Exception as e:
        logger.error(f"Unhandled error: {str(e)}")
        logger.error(traceback.format_exc())

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python test_problematic_file.py <excel_file_path> [sheet_name]")
        sys.exit(1)
    
    file_path = sys.argv[1]
    sheet_name = sys.argv[2] if len(sys.argv) > 2 else None
    
    test_file_processing(file_path, sheet_name)
