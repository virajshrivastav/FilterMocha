import pandas as pd
import os
import json
import logging
from datetime import datetime
import sys

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(f'excel_standardizer_{datetime.now().strftime("%Y%m%d")}.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class ExcelStandardizer:
    """Class to standardize Excel files to a specific format"""

    def __init__(self):
        """Initialize the standardizer"""
        self.log_entries = []
        self.standard_format_path = os.path.join(os.environ.get('STANDARD_FORMAT_DIR', 'Standard-Format'), 'iMocha Standard Format.xlsx')
        self.output_dir = os.environ.get('OUTPUT_FOLDER', 'Processed-Files')

        # Create output directory if it doesn't exist
        os.makedirs(self.output_dir, exist_ok=True)

    def analyze_file(self, file_path, sheet_name=None):
        """Analyze an Excel file and return column information"""
        try:
            # Get sheet names first
            xl = pd.ExcelFile(file_path)
            sheet_names = xl.sheet_names

            # Create a case-insensitive lookup dictionary for sheet names
            sheet_name_lookup = {name.lower().strip(): name for name in sheet_names}

            logger.info(f"Found sheets in file: {', '.join(sheet_names)}")

            # Collect information about all sheets
            sheet_info = {}
            for sheet in sheet_names:
                try:
                    # Try to read the sheet with different engines and options if one fails
                    try:
                        # First try with default settings
                        temp_df = pd.read_excel(file_path, sheet_name=sheet)
                    except Exception as first_error:
                        try:
                            # Try with openpyxl engine
                            logger.warning(f"First attempt to read sheet '{sheet}' failed, trying with openpyxl engine")
                            temp_df = pd.read_excel(file_path, sheet_name=sheet, engine='openpyxl')
                        except Exception:
                            try:
                                # Try with converters to handle mixed data types
                                logger.warning(f"Second attempt to read sheet '{sheet}' failed, trying with converters")
                                # Create a converter that converts everything to string
                                converters = {i: str for i in range(100)}  # Handle up to 100 columns
                                temp_df = pd.read_excel(file_path, sheet_name=sheet, converters=converters)
                            except Exception:
                                try:
                                    # Try with direct openpyxl access
                                    logger.warning(f"Third attempt to read sheet '{sheet}' failed, trying with direct openpyxl access")
                                    import openpyxl
                                    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
                                    ws = wb[sheet]
                                    data = []
                                    for row in ws.rows:
                                        data.append([cell.value for cell in row])
                                    if data:
                                        temp_df = pd.DataFrame(data[1:], columns=data[0])
                                    else:
                                        temp_df = pd.DataFrame()
                                except Exception:
                                    # If all attempts fail, raise the original error
                                    logger.error(f"All attempts to read sheet '{sheet}' failed")
                                    raise first_error

                    # Clean column names by stripping whitespace and handling special characters
                    original_columns = temp_df.columns.tolist()
                    cleaned_columns = []
                    for col in original_columns:
                        if isinstance(col, str):
                            # Strip whitespace and replace problematic characters
                            cleaned_col = col.strip()
                            # Replace non-breaking spaces with regular spaces
                            cleaned_col = cleaned_col.replace('\xa0', ' ').strip()
                            # Log if cleaning changed the column name
                            if cleaned_col != col:
                                logger.info(f"Cleaned column name: '{col}' -> '{cleaned_col}'")
                            cleaned_columns.append(cleaned_col)
                        else:
                            # Convert non-string columns to string
                            try:
                                cleaned_col = str(col).strip()
                                logger.info(f"Converted non-string column to string: {type(col).__name__} -> '{cleaned_col}'")
                                cleaned_columns.append(cleaned_col)
                            except Exception:
                                # If conversion fails, use a placeholder name
                                placeholder = f"Column_{len(cleaned_columns)}"
                                logger.warning(f"Could not convert column of type {type(col).__name__} to string, using placeholder: '{placeholder}'")
                                cleaned_columns.append(placeholder)

                    # Check for duplicate column names after cleaning
                    if len(cleaned_columns) != len(set(cleaned_columns)):
                        # Handle duplicate column names by adding suffixes
                        seen = {}
                        for i, col in enumerate(cleaned_columns):
                            if col in seen:
                                seen[col] += 1
                                cleaned_columns[i] = f"{col}_{seen[col]}"
                                logger.warning(f"Renamed duplicate column: '{col}' -> '{cleaned_columns[i]}'")
                            else:
                                seen[col] = 0

                    # Assign cleaned column names
                    temp_df.columns = cleaned_columns

                    # Convert all data to strings to handle mixed data types
                    for col in temp_df.columns:
                        try:
                            # Convert column to string, handling NaN values
                            temp_df[col] = temp_df[col].astype(str)
                            # Replace 'nan' strings with empty strings
                            temp_df[col] = temp_df[col].replace('nan', '')
                            # Replace non-breaking spaces with regular spaces
                            if temp_df[col].dtype == 'object':
                                temp_df[col] = temp_df[col].str.replace('\xa0', ' ')
                        except Exception as e:
                            logger.warning(f"Error converting column '{col}' to string: {str(e)}")

                    sheet_info[sheet] = {
                        'columns': len(temp_df.columns),
                        'rows': len(temp_df),
                        'df': temp_df
                    }
                except Exception as e:
                    # Log the error but continue with other sheets
                    logger.warning(f"Could not read sheet '{sheet}': {str(e)}")
                    # Add a placeholder for this sheet
                    sheet_info[sheet] = {
                        'columns': 0,
                        'rows': 0,
                        'df': pd.DataFrame(),
                        'error': str(e)
                    }

            # If a specific sheet is requested, use it (with case-insensitive matching)
            if sheet_name is not None:
                # Try exact match first
                if sheet_name in sheet_names:
                    actual_sheet_name = sheet_name
                    logger.info(f"Using exact sheet name match: '{sheet_name}'")
                # Then try case-insensitive match
                elif sheet_name.lower().strip() in sheet_name_lookup:
                    actual_sheet_name = sheet_name_lookup[sheet_name.lower().strip()]
                    logger.info(f"Using case-insensitive match: '{sheet_name}' -> '{actual_sheet_name}'")
                # Try fuzzy matching if no exact or case-insensitive match
                else:
                    # Find the closest match based on similarity
                    best_match = None
                    best_score = 0
                    for name in sheet_names:
                        # Calculate similarity score (simple implementation)
                        name_lower = name.lower()
                        sheet_name_lower = sheet_name.lower()

                        # Check if one is a substring of the other
                        if name_lower in sheet_name_lower or sheet_name_lower in name_lower:
                            score = 0.8  # High score for substring match
                        else:
                            # Count matching characters
                            common_chars = set(name_lower) & set(sheet_name_lower)
                            score = len(common_chars) / max(len(name_lower), len(sheet_name_lower))

                        if score > best_score:
                            best_score = score
                            best_match = name

                    if best_score > 0.5:  # Threshold for accepting a fuzzy match
                        actual_sheet_name = best_match
                        logger.info(f"Using fuzzy match: '{sheet_name}' -> '{actual_sheet_name}' (score: {best_score:.2f})")
                    else:
                        # If no good match found, raise a more descriptive error
                        available_sheets = ', '.join(sheet_names)
                        raise ValueError(f"Requested sheet '{sheet_name}' not found in the Excel file. Available sheets: {available_sheets}")

                # Check if this sheet had an error during loading
                if 'error' in sheet_info[actual_sheet_name]:
                    error_msg = sheet_info[actual_sheet_name]['error']
                    raise ValueError(f"Error reading sheet '{actual_sheet_name}': {error_msg}")
                df = sheet_info[actual_sheet_name]['df']
                sheet_name = actual_sheet_name  # Update sheet_name to the actual name for consistency
            # Otherwise, try to find the best sheet
            else:
                # First, check if 'Sheet1' exists - it's often the main data sheet
                if 'Sheet1' in sheet_names:
                    # Skip if Sheet1 had an error
                    if 'error' in sheet_info['Sheet1']:
                        # If Sheet1 has an error, try other sheets
                        best_sheet = None
                        max_columns = 0
                    else:
                        temp_df = sheet_info['Sheet1']['df']
                        # Verify it has data
                        if len(temp_df.columns) > 0 and len(temp_df) > 0:
                            sheet_name = 'Sheet1'
                            df = temp_df
                        else:
                            # If Sheet1 exists but is empty, try other sheets
                            best_sheet = None
                            max_columns = 0
                else:
                    # If Sheet1 doesn't exist, try to find the best sheet
                    best_sheet = None
                    max_columns = 0

                # If we haven't found a suitable sheet yet, check all sheets
                if sheet_name is None:
                    for sheet, info in sheet_info.items():
                        # Skip sheets with errors
                        if 'error' in info:
                            continue
                        # Prefer sheets with more columns and non-empty rows
                        if info['columns'] > max_columns and info['rows'] > 0:
                            max_columns = info['columns']
                            best_sheet = sheet

                    # If we found a good sheet, use it
                    if best_sheet:
                        sheet_name = best_sheet
                        df = sheet_info[best_sheet]['df']
                    else:
                        # Default to first sheet if we couldn't determine the best one
                        sheet_name = sheet_names[0]
                        df = sheet_info[sheet_name]['df']

            # Check if we found a valid sheet
            if df is None or len(df.columns) == 0:
                # No valid sheet found, raise an error
                raise ValueError("Could not find a valid sheet in the Excel file. All sheets are either empty or have errors.")

            # Get column information
            columns_info = []
            for col in df.columns:
                col_info = {
                    'name': col,
                    'type': str(df[col].dtype),
                    'sample_values': df[col].dropna().head(3).tolist()
                }
                columns_info.append(col_info)

            # Log the analysis
            self.log_entries.append({
                'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'action': 'File Analysis',
                'details': f'Analyzed file {os.path.basename(file_path)} (sheet: {sheet_name}) with {len(df)} rows and {len(df.columns)} columns'
            })

            # Remove the dataframe from sheet_info to make it JSON serializable
            sheet_info_clean = {}
            for sheet, info in sheet_info.items():
                sheet_info_clean[sheet] = {
                    'columns': info['columns'],
                    'rows': info['rows']
                }

            return columns_info, df.shape, sheet_names, sheet_name, sheet_info_clean
        except Exception as e:
            logger.error(f"Error analyzing file: {str(e)}")
            self.log_entries.append({
                'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'action': 'Error',
                'details': f'Failed to analyze file: {str(e)}'
            })
            raise

    def get_standard_columns(self):
        """Get the standard format columns"""
        try:
            standard_df = pd.read_excel(self.standard_format_path)
            return standard_df.columns.tolist()
        except Exception as e:
            logger.error(f"Error reading standard format: {str(e)}")
            self.log_entries.append({
                'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'action': 'Error',
                'details': f'Failed to read standard format: {str(e)}'
            })
            raise

    def process_file(self, input_file, mapping_config, split_config=None, custom_values=None, sheet_name=None):
        """Process an Excel file with the given mapping configuration"""
        try:
            # Use the analyze_file method to get sheet information and handle errors
            _, _, _, selected_sheet, _ = self.analyze_file(input_file, sheet_name)

            # Use the selected sheet from analyze_file
            sheet_name = selected_sheet
            logger.info(f"Using sheet '{sheet_name}' for processing")

            # Load the Excel file directly with the selected sheet using the same robust approach as in analyze_file
            try:
                # Try different methods to load the sheet
                try:
                    # First try with default settings
                    df = pd.read_excel(input_file, sheet_name=sheet_name)
                except Exception as first_error:
                    try:
                        # Try with openpyxl engine
                        logger.warning(f"First attempt to read sheet '{sheet_name}' failed, trying with openpyxl engine")
                        df = pd.read_excel(input_file, sheet_name=sheet_name, engine='openpyxl')
                    except Exception:
                        try:
                            # Try with converters to handle mixed data types
                            logger.warning(f"Second attempt to read sheet '{sheet_name}' failed, trying with converters")
                            # Create a converter that converts everything to string
                            converters = {i: str for i in range(100)}  # Handle up to 100 columns
                            df = pd.read_excel(input_file, sheet_name=sheet_name, converters=converters)
                        except Exception:
                            try:
                                # Try with direct openpyxl access
                                logger.warning(f"Third attempt to read sheet '{sheet_name}' failed, trying with direct openpyxl access")
                                import openpyxl
                                wb = openpyxl.load_workbook(input_file, read_only=True, data_only=True)
                                ws = wb[sheet_name]
                                data = []
                                for row in ws.rows:
                                    data.append([cell.value for cell in row])
                                if data:
                                    df = pd.DataFrame(data[1:], columns=data[0])
                                else:
                                    df = pd.DataFrame()
                            except Exception:
                                # If all attempts fail, raise the original error
                                logger.error(f"All attempts to read sheet '{sheet_name}' failed")
                                raise first_error

                # Clean column names by stripping whitespace and handling special characters
                original_columns = df.columns.tolist()
                cleaned_columns = []
                for col in original_columns:
                    if isinstance(col, str):
                        # Strip whitespace and replace problematic characters
                        cleaned_col = col.strip()
                        # Replace non-breaking spaces with regular spaces
                        cleaned_col = cleaned_col.replace('\xa0', ' ').strip()
                        # Log if cleaning changed the column name
                        if cleaned_col != col:
                            logger.info(f"Cleaned column name: '{col}' -> '{cleaned_col}'")
                        cleaned_columns.append(cleaned_col)
                    else:
                        # Convert non-string columns to string
                        try:
                            cleaned_col = str(col).strip()
                            logger.info(f"Converted non-string column to string: {type(col).__name__} -> '{cleaned_col}'")
                            cleaned_columns.append(cleaned_col)
                        except Exception:
                            # If conversion fails, use a placeholder name
                            placeholder = f"Column_{len(cleaned_columns)}"
                            logger.warning(f"Could not convert column of type {type(col).__name__} to string, using placeholder: '{placeholder}'")
                            cleaned_columns.append(placeholder)

                # Check for duplicate column names after cleaning
                if len(cleaned_columns) != len(set(cleaned_columns)):
                    # Handle duplicate column names by adding suffixes
                    seen = {}
                    for i, col in enumerate(cleaned_columns):
                        if col in seen:
                            seen[col] += 1
                            cleaned_columns[i] = f"{col}_{seen[col]}"
                            logger.warning(f"Renamed duplicate column: '{col}' -> '{cleaned_columns[i]}'")
                        else:
                            seen[col] = 0

                # Assign cleaned column names
                df.columns = cleaned_columns

                # Convert all data to strings to handle mixed data types
                for col in df.columns:
                    try:
                        # Convert column to string, handling NaN values
                        df[col] = df[col].astype(str)
                        # Replace 'nan' strings with empty strings
                        df[col] = df[col].replace('nan', '')
                        # Replace non-breaking spaces with regular spaces
                        if df[col].dtype == 'object':
                            df[col] = df[col].str.replace('\xa0', ' ')
                    except Exception as e:
                        logger.warning(f"Error converting column '{col}' to string: {str(e)}")

                # Log column information
                logger.info(f"Processing file with {len(df)} rows and {len(df.columns)} columns")
                logger.info(f"Columns: {df.columns.tolist()}")
            except Exception as e:
                logger.error(f"Error loading sheet '{sheet_name}': {str(e)}")
                raise ValueError(f"Error loading sheet '{sheet_name}': {str(e)}")

            # Create a folder for output files based on the original filename
            input_filename = os.path.basename(input_file)
            input_name = os.path.splitext(input_filename)[0]  # Get filename without extension
            output_folder = os.path.join(self.output_dir, input_name)

            # Create the output folder if it doesn't exist
            os.makedirs(output_folder, exist_ok=True)

            # Log the file loading
            self.log_entries.append({
                'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'action': 'File Loaded',
                'details': f'Loaded file {input_filename} with {len(df)} rows and {len(df.columns)} columns'
            })

            # Get standard columns
            standard_columns = self.get_standard_columns()

            # Create result dataframe with standard columns
            result_df = pd.DataFrame(columns=standard_columns)

            # Track errors and error questions
            errors = []
            error_questions = []

            # Apply mapping
            for std_col in standard_columns:
                # Check if this column has a custom value
                if custom_values and std_col in custom_values and custom_values[std_col]:
                    # Use the custom value for all rows
                    result_df[std_col] = custom_values[std_col]
                    continue

                # Check if this column is mapped
                input_col = mapping_config.get(std_col)

                # Create a case-insensitive column lookup to handle columns with trailing spaces
                column_lookup = {col.lower().strip().replace('\xa0', ' ') if isinstance(col, str) else str(col).lower().strip(): col for col in df.columns}

                # Try exact match first, then try with stripped whitespace
                if input_col and input_col in df.columns:
                    # Use the column as is
                    actual_input_col = input_col
                    logger.info(f"Using exact column match: '{input_col}'")
                elif input_col and isinstance(input_col, str) and input_col.lower().strip().replace('\xa0', ' ') in column_lookup:
                    # Use the actual column name from the lookup
                    actual_input_col = column_lookup[input_col.lower().strip().replace('\xa0', ' ')]
                    logger.info(f"Using case-insensitive column match: '{input_col}' -> '{actual_input_col}'")
                # Try fuzzy matching for column names
                elif input_col and isinstance(input_col, str):
                    # Find the closest match based on similarity
                    best_match = None
                    best_score = 0
                    input_col_lower = input_col.lower().strip().replace('\xa0', ' ')

                    # Try common variations first
                    common_variations = {
                        'question type': ['q type', 'qtype', 'type', 'question', 'q_type'],
                        'difficulty level': ['level', 'difficulty', 'diff level', 'diff', 'difficulty_level'],
                        'question text': ['q text', 'qtext', 'question', 'text', 'q_text', 'question_text'],
                        'option (a)': ['option a', 'option/ answer 1', 'option 1', 'option/answer 1', 'answer 1', 'a'],
                        'option (b)': ['option b', 'option/ answer 2', 'option 2', 'option/answer 2', 'answer 2', 'b'],
                        'option (c)': ['option c', 'option/ answer 3', 'option 3', 'option/answer 3', 'answer 3', 'c'],
                        'option (d)': ['option d', 'option/ answer 4', 'option 4', 'option/answer 4', 'answer 4', 'd'],
                        'option (e)': ['option e', 'option/ answer 5', 'option 5', 'option/answer 5', 'answer 5', 'e'],
                        'option (f)': ['option f', 'option/ answer 6', 'option 6', 'option/answer 6', 'answer 6', 'f'],
                        'correct answer': ['answer', 'correct', 'correct_answer', 'right answer', 'right_answer'],
                        'answer explanation': ['explanation', 'answer_explanation', 'solution', 'rationale'],
                        'score': ['marks', 'points', 'value', 'weight'],
                        'topics': ['topic', 'subject', 'category', 'skill', 'topics_list'],
                        'author': ['author name', 'created by', 'writer', 'author_name', 'author email', 'author\'s email']
                    }

                    # Check if input_col matches any common variation
                    for standard_col, variations in common_variations.items():
                        if input_col_lower == standard_col or input_col_lower in variations:
                            # Try to find this standard column or its variations in the dataframe
                            for col in df.columns:
                                col_lower = col.lower().strip().replace('\xa0', ' ')
                                if col_lower == standard_col or col_lower in variations:
                                    best_match = col
                                    best_score = 1.0  # Perfect match through variations
                                    logger.info(f"Found match through common variations: '{input_col}' -> '{best_match}'")
                                    break

                    # If no match found through common variations, try fuzzy matching
                    if best_match is None:
                        for col in df.columns:
                            if not isinstance(col, str):
                                col_str = str(col)
                            else:
                                col_str = col

                            col_lower = col_str.lower().strip().replace('\xa0', ' ')

                            # Check if one is a substring of the other
                            if col_lower in input_col_lower or input_col_lower in col_lower:
                                score = 0.8  # High score for substring match
                            else:
                                # Count matching characters
                                common_chars = set(col_lower) & set(input_col_lower)
                                score = len(common_chars) / max(len(col_lower), len(input_col_lower))

                            if score > best_score:
                                best_score = score
                                best_match = col

                    if best_score > 0.5:  # Lower threshold to catch more matches
                        actual_input_col = best_match
                        logger.info(f"Using fuzzy column match: '{input_col}' -> '{actual_input_col}' (score: {best_score:.2f})")
                    else:
                        # No good match found
                        actual_input_col = None
                        logger.warning(f"No match found for column '{input_col}'. Available columns: {df.columns.tolist()}")
                else:
                    # Column not found, skip to the else block below
                    actual_input_col = None

                if actual_input_col:
                    # Copy the column data
                    result_df[std_col] = df[actual_input_col]

                    # Apply standardization for specific columns
                    if std_col == 'Question Type':
                        # Get the question text column for error tracking
                        question_col = mapping_config.get('Question Text')

                        # Apply the same case-insensitive matching for the question column
                        if question_col and question_col in df.columns:
                            actual_question_col = question_col
                        elif question_col and isinstance(question_col, str) and question_col.lower().strip() in column_lookup:
                            actual_question_col = column_lookup[question_col.lower().strip()]
                            logger.info(f"Using case-insensitive match for question column: '{question_col}' -> '{actual_question_col}'")
                        else:
                            actual_question_col = None

                        question_texts = df[actual_question_col].tolist() if actual_question_col else [''] * len(df)

                        # Apply standardization with error tracking
                        for i, value in enumerate(result_df[std_col]):
                            std_value, has_error = self._standardize_question_type(value)
                            result_df.at[i, std_col] = std_value
                            if has_error:
                                errors.append(f"Unknown Question Type: '{value}'")
                                error_questions.append(f"Row {i+2}: {question_texts[i] if i < len(question_texts) else 'Unknown'}")

                    elif std_col == 'Difficulty Level':
                        # Get the question text column for error tracking
                        question_col = mapping_config.get('Question Text')

                        # Apply the same case-insensitive matching for the question column
                        if question_col and question_col in df.columns:
                            actual_question_col = question_col
                        elif question_col and isinstance(question_col, str) and question_col.lower().strip() in column_lookup:
                            actual_question_col = column_lookup[question_col.lower().strip()]
                            logger.info(f"Using case-insensitive match for question column: '{question_col}' -> '{actual_question_col}'")
                        else:
                            actual_question_col = None

                        question_texts = df[actual_question_col].tolist() if actual_question_col else [''] * len(df)

                        # Apply standardization with error tracking
                        for i, value in enumerate(result_df[std_col]):
                            std_value, has_error = self._standardize_difficulty_level(value)
                            result_df.at[i, std_col] = std_value
                            if has_error:
                                errors.append(f"Unknown Difficulty Level: '{value}'")
                                error_questions.append(f"Row {i+2}: {question_texts[i] if i < len(question_texts) else 'Unknown'}")

                    elif std_col == 'Correct Answer':
                        # Get the question text column for error tracking
                        question_col = mapping_config.get('Question Text')

                        # Apply the same case-insensitive matching for the question column
                        if question_col and question_col in df.columns:
                            actual_question_col = question_col
                        elif question_col and isinstance(question_col, str) and question_col.lower().strip() in column_lookup:
                            actual_question_col = column_lookup[question_col.lower().strip()]
                            logger.info(f"Using case-insensitive match for question column: '{question_col}' -> '{actual_question_col}'")
                        else:
                            actual_question_col = None

                        question_texts = df[actual_question_col].tolist() if actual_question_col else [''] * len(df)

                        # Apply standardization with error tracking
                        for i, value in enumerate(result_df[std_col]):
                            std_value, has_error = self._standardize_correct_answer(value)
                            result_df.at[i, std_col] = std_value
                            if has_error:
                                errors.append(f"Invalid Correct Answer format: '{value}'")
                                error_questions.append(f"Row {i+2}: {question_texts[i] if i < len(question_texts) else 'Unknown'}")
                else:
                    # Column not mapped or not found - ensure it exists but is empty
                    result_df[std_col] = None

                    # Only report errors for required fields
                    if std_col not in ['Recording Time Limit:(Upto 5 mins)',
                                      'Retake Allowed:(Upto 5 mins)',
                                      'Set Prep Time (0.5 to 5 mins)',
                                      'Proofreading Status',
                                      'Editor Email',
                                      'Differential Scoring',
                                      'Answer Explanation',
                                      'Topics']:  # Topics is no longer mandatory
                        errors.append(f"Required column '{std_col}' not mapped or not found")
                        error_questions.append("Row 0 (N/A)")  # Used when column is missing entirely

            # Log the mapping process
            self.log_entries.append({
                'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'action': 'Mapping Applied',
                'details': f'Applied mapping configuration'
            })

            # Process file splitting if configured
            output_files = []
            if split_config and split_config.get('column'):
                split_column = split_config['column']
                if split_column in df.columns:
                    # Get unique values in the split column
                    unique_values = df[split_column].dropna().unique()

                    for value in unique_values:
                        # Filter rows for this value
                        try:
                            # First try with exact matching
                            value_mask = df[split_column] == value
                            filtered_indices = df.index[value_mask]

                            # Check if we got any matches
                            if len(filtered_indices) == 0:
                                # Try case-insensitive matching for string values
                                if isinstance(value, str):
                                    # Convert column to string and do case-insensitive comparison
                                    value_mask = df[split_column].astype(str).str.lower() == value.lower()
                                    filtered_indices = df.index[value_mask]

                            # Create a copy of the filtered result dataframe
                            value_df = result_df.loc[filtered_indices].copy()

                            # If still no matches, log a warning but continue
                            if len(value_df) == 0:
                                logger.warning(f"No rows found for {split_column}='{value}'. This may indicate a data issue.")
                                # Create an empty dataframe with the same columns
                                value_df = pd.DataFrame(columns=result_df.columns)
                        except Exception as e:
                            logger.error(f"Error filtering rows for {split_column}='{value}': {str(e)}")
                            # Create an empty dataframe with the same columns as a fallback
                            value_df = pd.DataFrame(columns=result_df.columns)

                        # Create a safe filename
                        safe_value = str(value).replace('/', '_').replace('\\', '_')
                        safe_value = ''.join(c for c in safe_value if c.isalnum() or c in '._- ')

                        # Save to file in the output folder
                        output_filename = f"{safe_value}.xlsx"
                        output_path = os.path.join(output_folder, output_filename)
                        try:
                            # Ensure all columns are in the correct order according to standard format
                            # Make a copy to avoid modifying the original DataFrame
                            value_df_copy = value_df.copy()
                            # Check if all standard columns exist in the DataFrame
                            missing_columns = [col for col in standard_columns if col not in value_df_copy.columns]
                            if missing_columns:
                                # Add missing columns with empty values
                                for col in missing_columns:
                                    value_df_copy[col] = None
                            # Reorder columns according to standard format
                            value_df_copy = value_df_copy[standard_columns]

                            # Use xlsxwriter for better formatting control
                            try:
                                with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
                                    # First check if the dataframe is empty
                                    if len(value_df_copy) == 0:
                                        # Add at least one empty row to avoid errors
                                        empty_row = {col: '' for col in value_df_copy.columns}
                                        value_df_copy = pd.DataFrame([empty_row], columns=value_df_copy.columns)

                                    # Write the dataframe to Excel
                                    value_df_copy.to_excel(writer, index=False, sheet_name='Questions')

                                    # Get the xlsxwriter workbook and worksheet objects
                                    workbook = writer.book
                                    worksheet = writer.sheets['Questions']

                                    # Set column widths for better readability - increased widths for all columns
                                    worksheet.set_column('A:A', 20)  # Question Type
                                    worksheet.set_column('B:B', 20)  # Difficulty Level
                                    worksheet.set_column('C:C', 80)  # Question Text
                                    worksheet.set_column('D:G', 50)  # Options A-D
                                    worksheet.set_column('H:H', 20)  # Option E
                                    worksheet.set_column('I:I', 20)  # Correct Answer
                                    worksheet.set_column('J:J', 80)  # Answer Explanation
                                    worksheet.set_column('K:K', 15)  # Score
                                    worksheet.set_column('L:L', 40)  # Topics
                                    worksheet.set_column('M:Z', 30)  # Other columns

                                    # Add a header format
                                    header_format = workbook.add_format({
                                        'bold': True,
                                        'text_wrap': True,
                                        'valign': 'top',
                                        'border': 1
                                    })

                                    # Write the column headers with the defined format
                                    for col_num, col_value in enumerate(value_df_copy.columns.values):
                                        worksheet.write(0, col_num, col_value, header_format)
                            except Exception as e:
                                logger.error(f"Error using xlsxwriter: {str(e)}")
                                # Try with a simpler approach
                                try:
                                    # First check if the dataframe is empty
                                    if len(value_df_copy) == 0:
                                        # Add at least one empty row to avoid errors
                                        empty_row = {col: '' for col in value_df_copy.columns}
                                        value_df_copy = pd.DataFrame([empty_row], columns=value_df_copy.columns)

                                    # Try with openpyxl engine
                                    value_df_copy.to_excel(output_path, index=False, engine='openpyxl', sheet_name='Questions')
                                except Exception as e2:
                                    logger.error(f"Error using openpyxl: {str(e2)}")
                                    # Last resort - try with default engine
                                    value_df_copy.to_excel(output_path, index=False, sheet_name='Questions')
                        except Exception as e:
                            logger.error(f"Error saving file {output_filename}: {str(e)}")
                            # Try with a different approach as fallback
                            try:
                                # First check if the dataframe is empty
                                if len(value_df) == 0:
                                    # Add at least one empty row to avoid errors
                                    empty_row = {col: '' for col in result_df.columns}
                                    value_df = pd.DataFrame([empty_row], columns=result_df.columns)

                                # Try with openpyxl engine
                                value_df.to_excel(output_path, index=False, engine='openpyxl')
                            except Exception as e2:
                                logger.error(f"Fallback error with openpyxl: {str(e2)}")
                                try:
                                    # Last resort - try with default engine and minimal formatting
                                    value_df.to_excel(output_path, index=False)
                                except Exception as e3:
                                    logger.error(f"All attempts to save file failed: {str(e3)}")
                                    # Add to errors list
                                    errors.append(f"Failed to save file for {split_column}='{value}': {str(e3)}")
                        output_files.append(output_path)

                        self.log_entries.append({
                            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                            'action': 'File Split',
                            'details': f'Created split file for {split_column}="{value}" with {len(value_df)} rows'
                        })
                else:
                    errors.append(f"Split column '{split_column}' not found in input file")
            else:
                # Save the entire result to a single file in the output folder
                # Keep the same name format for non-split files
                output_filename = f"processed_{os.path.basename(input_file)}"
                output_path = os.path.join(output_folder, output_filename)
                try:
                    # Ensure all columns are in the correct order according to standard format
                    # Make a copy to avoid modifying the original DataFrame
                    result_df_copy = result_df.copy()
                    # Check if all standard columns exist in the DataFrame
                    missing_columns = [col for col in standard_columns if col not in result_df_copy.columns]
                    if missing_columns:
                        # Add missing columns with empty values
                        for col in missing_columns:
                            result_df_copy[col] = None
                    # Reorder columns according to standard format
                    result_df_copy = result_df_copy[standard_columns]

                    # Use xlsxwriter for better formatting control
                    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
                        result_df_copy.to_excel(writer, index=False, sheet_name='Questions')

                        # Get the xlsxwriter workbook and worksheet objects
                        workbook = writer.book
                        worksheet = writer.sheets['Questions']

                        # Set column widths for better readability - increased widths for all columns
                        worksheet.set_column('A:A', 20)  # Question Type
                        worksheet.set_column('B:B', 20)  # Difficulty Level
                        worksheet.set_column('C:C', 80)  # Question Text
                        worksheet.set_column('D:G', 50)  # Options A-D
                        worksheet.set_column('H:H', 20)  # Option E
                        worksheet.set_column('I:I', 20)  # Correct Answer
                        worksheet.set_column('J:J', 80)  # Answer Explanation
                        worksheet.set_column('K:K', 15)  # Score
                        worksheet.set_column('L:L', 40)  # Topics
                        worksheet.set_column('M:Z', 30)  # Other columns

                        # Add a header format
                        header_format = workbook.add_format({
                            'bold': True,
                            'text_wrap': True,
                            'valign': 'top',
                            'border': 1
                        })

                        # Write the column headers with the defined format
                        for col_num, value in enumerate(result_df_copy.columns.values):
                            worksheet.write(0, col_num, value, header_format)
                except Exception as e:
                    import traceback
                    error_traceback = traceback.format_exc()
                    logger.error(f"Error saving file {output_filename}: {str(e)}\n{error_traceback}")
                    # Try with a different engine as fallback
                    try:
                        result_df_copy.to_excel(output_path, index=False, engine='openpyxl')
                    except Exception as e2:
                        logger.error(f"Second attempt failed: {str(e2)}")
                        try:
                            # Last resort - basic Excel output
                            result_df.to_excel(output_path, index=False)
                        except Exception as e3:
                            logger.error(f"All attempts failed: {str(e3)}")
                            raise Exception(f"Could not save file {output_filename} after multiple attempts")
                output_files.append(output_path)

                self.log_entries.append({
                    'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'action': 'File Saved',
                    'details': f'Saved processed file with {len(result_df)} rows'
                })

            # Save log file in the output folder
            log_filename = f"log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            log_path = os.path.join(output_folder, log_filename)
            with open(log_path, 'w') as f:
                json.dump({
                    'file_name': os.path.basename(input_file),
                    'processing_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'entries': self.log_entries,
                    'errors': errors
                }, f, indent=2)

            # Apply custom values if provided
            if custom_values:
                for column, value in custom_values.items():
                    if column in result_df.columns:
                        result_df[column] = value
                        self.log_entries.append({
                            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                            'action': 'Custom Value Applied',
                            'details': f'Applied custom value "{value}" to column "{column}"'
                        })

            # Return the result
            return {
                'output_files': output_files,
                'log_file': log_path,
                'errors': errors,
                'error_questions': error_questions
            }
        except Exception as e:
            logger.error(f"Error processing file: {str(e)}")
            self.log_entries.append({
                'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'action': 'Error',
                'details': f'Failed to process file: {str(e)}'
            })
            raise

    def _standardize_question_type(self, value):
        """Standardize Question Type values"""
        if pd.isna(value):
            return None, True

        value = str(value).strip().lower()

        if value in ['mcq', 'single', 'one answer', 'single choice', 'single select']:
            return 'MCQ', False
        elif value in ['maq', 'multiple', 'multiple answers', 'multiple choice', 'multi select']:
            return 'MAQ', False
        elif value in ['true/false', 'true or false', 'yes or no', 't/f', 'yes/no']:
            return 'True/False', False
        elif value in ['fib', 'fill in the blank', 'fill in blank', 'fill blank', 'fill-in-the-blank']:
            return 'FIB', False
        elif value in ['desc', 'descriptive', 'long answer', 'essay', 'paragraph', 'long', 'descriptive question']:
            return 'DESC', False
        else:
            return value, True

    def _standardize_difficulty_level(self, value):
        """Standardize Difficulty Level values"""
        if pd.isna(value):
            return None, True

        value = str(value).strip().lower()

        if value in ['easy', 'beginner', 'basic', 'e']:
            return 'Easy', False
        elif value in ['medium', 'intermediate', 'moderate', 'm']:
            return 'Medium', False
        elif value in ['hard', 'difficult', 'advanced', 'expert', 'h']:
            return 'Hard', False
        else:
            return value, True

    def _standardize_correct_answer(self, value):
        """Standardize Correct Answer values"""
        if pd.isna(value):
            return None, True

        value = str(value).strip().lower()

        # Convert numeric answers to alphabetic (1->a, 2->b, etc.)
        if value.isdigit():
            num = int(value)
            if 1 <= num <= 26:
                return chr(96 + num), False  # ASCII: 'a' is 97, so 1 -> 'a', 2 -> 'b', etc.
            else:
                return value, True

        # Handle comma-separated values for multiple answers
        if ',' in value:
            parts = [part.strip() for part in value.split(',')]
            standardized_parts = []
            has_error = False

            for part in parts:
                if part.isdigit():
                    num = int(part)
                    if 1 <= num <= 26:
                        standardized_parts.append(chr(96 + num))
                    else:
                        standardized_parts.append(part)
                        has_error = True
                else:
                    standardized_parts.append(part)

            return ','.join(standardized_parts), has_error

        return value, False

def main():
    """Command-line interface for the Excel Standardizer"""
    if len(sys.argv) < 2:
        print("Usage: python excel_standardizer.py <input_file> [mapping_file]")
        sys.exit(1)

    input_file = sys.argv[1]
    mapping_file = sys.argv[2] if len(sys.argv) > 2 else None

    standardizer = ExcelStandardizer()

    try:
        # If mapping file is provided, use it
        if mapping_file:
            with open(mapping_file, 'r') as f:
                mapping_config = json.load(f)
        else:
            # Otherwise, try to auto-map columns
            columns_info, _ = standardizer.analyze_file(input_file)
            standard_columns = standardizer.get_standard_columns()

            mapping_config = {}
            for std_col in standard_columns:
                for col_info in columns_info:
                    col_name = col_info['name']
                    if (col_name.lower() in std_col.lower() or
                        std_col.lower() in col_name.lower()):
                        mapping_config[std_col] = col_name
                        break

        # Process the file
        result = standardizer.process_file(input_file, mapping_config)

        # Print results
        print(f"\nProcessing complete!")
        print(f"Output files ({len(result['output_files'])}):")
        for file in result['output_files']:
            print(f"- {file}")

        if result['errors']:
            print(f"\nWarnings/Errors ({len(result['errors'])}):")
            for error in result['errors'][:10]:  # Show first 10 errors
                print(f"- {error}")

            if len(result['errors']) > 10:
                print(f"... and {len(result['errors']) - 10} more")

        print(f"\nLog file: {result['log_file']}")

    except Exception as e:
        print(f"Error: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()
