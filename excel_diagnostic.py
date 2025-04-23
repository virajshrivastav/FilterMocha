import pandas as pd
import os
import sys
import logging
import traceback

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler()
    ]
)
logger = logging.getLogger()

def diagnose_excel_file(file_path):
    """Perform a comprehensive diagnosis of an Excel file"""
    logger.info(f"Diagnosing Excel file: {file_path}")
    
    if not os.path.exists(file_path):
        logger.error(f"File not found: {file_path}")
        return
    
    try:
        # Try to open the Excel file
        logger.info("Attempting to open the Excel file...")
        xl = pd.ExcelFile(file_path)
        
        # Get sheet names
        sheet_names = xl.sheet_names
        logger.info(f"Found {len(sheet_names)} sheets: {', '.join(sheet_names)}")
        
        # Create a case-insensitive lookup for sheet names
        sheet_name_lookup = {name.lower(): name for name in sheet_names}
        logger.info(f"Case-insensitive sheet lookup: {sheet_name_lookup}")
        
        # Examine each sheet
        for sheet_name in sheet_names:
            logger.info(f"\nExamining sheet: {sheet_name}")
            
            try:
                # Try to read the sheet
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                logger.info(f"  Successfully read sheet with {len(df)} rows and {len(df.columns)} columns")
                
                # Check column names
                logger.info(f"  Column names:")
                for i, col in enumerate(df.columns):
                    col_type = type(col).__name__
                    col_repr = repr(col)
                    col_str = str(col)
                    logger.info(f"    {i+1}. {col_repr} (type: {col_type})")
                    
                    # Check for whitespace or special characters
                    if isinstance(col, str):
                        if col != col.strip():
                            logger.warning(f"      Column has leading/trailing whitespace: '{col}' vs '{col.strip()}'")
                        if '\n' in col or '\r' in col or '\t' in col:
                            logger.warning(f"      Column contains special characters")
                
                # Check for duplicate column names
                if len(df.columns) != len(set(df.columns)):
                    logger.warning("  Sheet contains duplicate column names!")
                    col_counts = {}
                    for col in df.columns:
                        col_counts[col] = col_counts.get(col, 0) + 1
                    for col, count in col_counts.items():
                        if count > 1:
                            logger.warning(f"    Column '{col}' appears {count} times")
                
                # Check data types
                logger.info(f"  Column data types:")
                for col, dtype in df.dtypes.items():
                    logger.info(f"    {col}: {dtype}")
                
                # Check for NaN values
                nan_counts = df.isna().sum()
                if nan_counts.sum() > 0:
                    logger.info(f"  NaN value counts:")
                    for col, count in nan_counts.items():
                        if count > 0:
                            logger.info(f"    {col}: {count} NaN values")
                
            except Exception as e:
                logger.error(f"  Error reading sheet '{sheet_name}': {str(e)}")
                logger.error(f"  {traceback.format_exc()}")
    
    except Exception as e:
        logger.error(f"Error diagnosing Excel file: {str(e)}")
        logger.error(f"{traceback.format_exc()}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python excel_diagnostic.py <excel_file_path>")
        sys.exit(1)
    
    file_path = sys.argv[1]
    diagnose_excel_file(file_path)
