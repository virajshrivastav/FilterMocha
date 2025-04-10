import pandas as pd
import sys
import os
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

def analyze_excel_file(file_path):
    """Analyze an Excel file and print detailed information about its structure"""
    try:
        logger.info(f"Analyzing Excel file: {file_path}")
        
        # Check if the file exists
        if not os.path.exists(file_path):
            logger.error(f"File not found: {file_path}")
            return
        
        # Try to open the Excel file
        try:
            xl = pd.ExcelFile(file_path)
            sheet_names = xl.sheet_names
            logger.info(f"Successfully opened file. Found {len(sheet_names)} sheets: {sheet_names}")
            
            # Analyze each sheet
            for sheet_name in sheet_names:
                logger.info(f"\nAnalyzing sheet: {sheet_name}")
                
                try:
                    # Try to read the sheet
                    df = pd.read_excel(file_path, sheet_name=sheet_name)
                    logger.info(f"  Successfully read sheet with {len(df)} rows and {len(df.columns)} columns")
                    
                    # Check column names
                    logger.info(f"  Column names:")
                    for i, col in enumerate(df.columns):
                        col_type = type(col).__name__
                        col_repr = repr(col)
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
                    
                    # Check for problematic values
                    logger.info(f"  Checking for problematic values...")
                    for col in df.columns:
                        try:
                            # Check for mixed data types
                            if df[col].dtype == 'object':
                                types = df[col].apply(lambda x: type(x).__name__ if not pd.isna(x) else 'nan').unique()
                                if len(types) > 1 and 'nan' in types:
                                    types = [t for t in types if t != 'nan']
                                if len(types) > 1:
                                    logger.warning(f"    Column '{col}' has mixed data types: {', '.join(types)}")
                            
                            # Check for extremely long values
                            if df[col].dtype == 'object':
                                max_len = df[col].astype(str).apply(len).max()
                                if max_len > 1000:
                                    logger.warning(f"    Column '{col}' has very long values (max length: {max_len})")
                                    
                            # Check for special characters
                            if df[col].dtype == 'object':
                                special_chars = df[col].astype(str).str.contains('[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\xFF]', regex=True).sum()
                                if special_chars > 0:
                                    logger.warning(f"    Column '{col}' has {special_chars} values with special characters")
                        except Exception as e:
                            logger.warning(f"    Error checking column '{col}': {str(e)}")
                    
                    # Sample data
                    logger.info(f"  Sample data (first 2 rows):")
                    try:
                        sample = df.head(2)
                        for i, row in sample.iterrows():
                            logger.info(f"    Row {i+1}:")
                            for col in sample.columns:
                                val = row[col]
                                val_type = type(val).__name__
                                val_repr = repr(val)
                                logger.info(f"      {col}: {val_repr} (type: {val_type})")
                    except Exception as e:
                        logger.warning(f"    Error displaying sample data: {str(e)}")
                    
                except Exception as e:
                    logger.error(f"  Error reading sheet '{sheet_name}': {str(e)}")
                    logger.error(traceback.format_exc())
        
        except Exception as e:
            logger.error(f"Error opening Excel file: {str(e)}")
            logger.error(traceback.format_exc())
    
    except Exception as e:
        logger.error(f"Error analyzing Excel file: {str(e)}")
        logger.error(traceback.format_exc())

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python analyze_excel.py <excel_file_path>")
        sys.exit(1)
    
    file_path = sys.argv[1]
    analyze_excel_file(file_path)
