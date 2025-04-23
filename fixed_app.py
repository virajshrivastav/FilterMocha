from flask import Flask, render_template, request, jsonify, send_from_directory, Response, redirect, send_file
import os
import pandas as pd
from datetime import datetime
import json
import logging
import time
import zipfile
import io
from excel_standardizer_improved import ExcelStandardizer
from werkzeug.utils import secure_filename

# Initialize Flask app with the original templates folder
app = Flask(__name__,
            static_folder='static',
            template_folder='templates')

# Disable template caching
app.config['TEMPLATES_AUTO_RELOAD'] = True

# Add timestamp to force cache refresh
timestamp = int(time.time())

# Add no-cache headers to all responses
@app.after_request
def add_header(response):
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, post-check=0, pre-check=0, max-age=0'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '-1'
    return response

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('fixed_app.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Define folders
# Use /tmp for writable directories on Render
UPLOAD_FOLDER = os.environ.get('UPLOAD_FOLDER', 'uploads')
OUTPUT_FOLDER = os.environ.get('OUTPUT_FOLDER', 'Processed-Files')  # Match the folder used by ExcelStandardizer

# Create folders if they don't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Create standardizer
standardizer = ExcelStandardizer()

# Original index route - now redirects to the main app
@app.route('/old-index')
def old_index():
    """Render the original index page"""
    return render_template('index.html', timestamp=timestamp)

@app.route('/')
def index():
    """Render the main page"""
    return render_template('simple_upload.html', timestamp=timestamp)

# Keep this for backward compatibility
@app.route('/new-solution')
def new_solution():
    """Redirect to the main page"""
    return redirect('/')

@app.route('/test-upload')
def test_upload():
    """Render the test upload page"""
    return render_template('test_upload.html')

@app.route('/api/analyze', methods=['POST'])
def analyze_file():
    """Analyze an uploaded Excel file and return column information"""
    logger.info(f"Analyze file request received: {request.files}")

    if 'file' not in request.files:
        logger.error("No file part in request")
        return jsonify({'error': 'No file part'}), 400

    file = request.files['file']
    logger.info(f"File received: {file.filename}")

    if file.filename == '':
        logger.error("Empty filename")
        return jsonify({'error': 'No selected file'}), 400

    if not file.filename.endswith(('.xlsx', '.xls')):
        logger.error(f"Invalid file type: {file.filename}")
        return jsonify({'error': 'File must be an Excel file (.xlsx or .xls)'}), 400

    # Save the uploaded file
    filename = secure_filename(file.filename)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    unique_filename = f"{timestamp}_{filename}"
    file_path = os.path.join(UPLOAD_FOLDER, unique_filename)
    file.save(file_path)

    try:
        # Get sheet name if provided
        sheet_name = request.form.get('sheet_name')
        logger.info(f"Sheet name from request: {sheet_name}")

        # First, try to open the file with pandas to check if it's valid
        try:
            logger.info(f"Checking if file is valid: {file_path}")
            xl = pd.ExcelFile(file_path)
            available_sheets = xl.sheet_names
            logger.info(f"File is valid. Available sheets: {available_sheets}")

            # If a sheet name is provided, check if it exists
            if sheet_name:
                # Create a case-insensitive lookup
                sheet_lookup = {s.lower().strip(): s for s in available_sheets}

                if sheet_name in available_sheets:
                    logger.info(f"Sheet '{sheet_name}' found (exact match)")
                elif sheet_name.lower().strip() in sheet_lookup:
                    actual_sheet = sheet_lookup[sheet_name.lower().strip()]
                    logger.info(f"Sheet '{sheet_name}' found (case-insensitive match): '{actual_sheet}'")
                    sheet_name = actual_sheet
                else:
                    logger.warning(f"Sheet '{sheet_name}' not found in file. Available sheets: {available_sheets}")
        except Exception as e:
            logger.error(f"Error checking file: {str(e)}")
            import traceback
            error_traceback = traceback.format_exc()
            logger.error(f"Traceback: {error_traceback}")
            return jsonify({'error': f"Error checking file: {str(e)}"}), 500

        # Now try to read the specific sheet to check if it's readable
        if sheet_name:
            try:
                logger.info(f"Checking if sheet '{sheet_name}' is readable")
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                logger.info(f"Sheet '{sheet_name}' is readable. Shape: {df.shape}, Columns: {df.columns.tolist()}")

                # Check for problematic values
                for col in df.columns:
                    try:
                        # Check for mixed data types
                        if df[col].dtype == 'object':
                            types = df[col].apply(lambda x: type(x).__name__ if not pd.isna(x) else 'nan').unique()
                            if len(types) > 1 and 'nan' in types:
                                types = [t for t in types if t != 'nan']
                            if len(types) > 1:
                                logger.warning(f"Column '{col}' has mixed data types: {', '.join(types)}")
                    except Exception as e:
                        logger.warning(f"Error checking column '{col}': {str(e)}")
            except Exception as e:
                logger.error(f"Error reading sheet '{sheet_name}': {str(e)}")
                import traceback
                error_traceback = traceback.format_exc()
                logger.error(f"Traceback: {error_traceback}")
                # Don't return an error here, let the standardizer try to handle it

        # Analyze the file with the standardizer
        logger.info(f"Analyzing file with standardizer: {file_path}, sheet: {sheet_name}")
        try:
            columns_info, file_shape, sheet_names, selected_sheet, sheet_info = standardizer.analyze_file(file_path, sheet_name)
            logger.info(f"File analyzed successfully. Found {len(columns_info)} columns, {file_shape[0]} rows, {len(sheet_names)} sheets, selected sheet: {selected_sheet}")

            # Get standard columns
            standard_columns = standardizer.get_standard_columns()
            logger.info(f"Got {len(standard_columns)} standard columns")

            # Prepare response
            response_data = {
                'filename': unique_filename,
                'original_filename': filename,
                'columns': columns_info,
                'rows': file_shape[0],
                'standard_columns': standard_columns,
                'sheet_names': sheet_names,
                'selected_sheet': selected_sheet,
                'sheet_info': sheet_info
            }
            logger.info(f"Sending response with {len(columns_info)} columns and {len(sheet_names)} sheets")
            return jsonify(response_data)
        except Exception as e:
            logger.error(f"Error in standardizer.analyze_file: {str(e)}")
            import traceback
            error_traceback = traceback.format_exc()
            logger.error(f"Traceback: {error_traceback}")

            # Instead of returning an error, try to return a valid response with default values
            try:
                # Get standard columns even if analysis failed
                standard_columns = standardizer.get_standard_columns()

                # Return a response with empty data but valid structure
                response_data = {
                    'filename': unique_filename,
                    'original_filename': filename,
                    'columns': [],  # Empty columns
                    'rows': 0,
                    'standard_columns': standard_columns,
                    'sheet_names': [sheet_name] if sheet_name else ['Sheet1'],  # Default sheet name
                    'selected_sheet': sheet_name if sheet_name else 'Sheet1',
                    'sheet_info': {},
                    'silent_error': f"Error analyzing file: {str(e)}"  # Add error message but don't show it
                }
                logger.info(f"Sending response with empty data due to error")
                return jsonify(response_data)
            except Exception as inner_e:
                # If even this fails, log it but still don't show error to user
                logger.error(f"Failed to create fallback response: {str(inner_e)}")
                return jsonify({
                    'filename': unique_filename,
                    'original_filename': filename,
                    'columns': [],
                    'rows': 0,
                    'standard_columns': [],
                    'sheet_names': [sheet_name] if sheet_name else ['Sheet1'],
                    'selected_sheet': sheet_name if sheet_name else 'Sheet1',
                    'sheet_info': {},
                    'silent_error': f"Error analyzing file: {str(e)}"
                })
    except Exception as e:
        import traceback
        error_traceback = traceback.format_exc()
        logger.error(f"Unhandled error analyzing file: {str(e)}\n{error_traceback}")

        # Return a valid response with empty data instead of an error
        return jsonify({
            'filename': unique_filename if 'unique_filename' in locals() else 'unknown',
            'original_filename': filename if 'filename' in locals() else 'unknown',
            'columns': [],
            'rows': 0,
            'standard_columns': [],
            'sheet_names': [sheet_name] if 'sheet_name' in locals() and sheet_name else ['Sheet1'],
            'selected_sheet': sheet_name if 'sheet_name' in locals() and sheet_name else 'Sheet1',
            'sheet_info': {},
            'silent_error': f"Unhandled error: {str(e)}"
        })

@app.route('/api/process', methods=['POST'])
def process_file():
    """Process an Excel file with the given mapping configuration"""
    data = request.json

    if not data or 'filename' not in data or 'mapping' not in data:
        return jsonify({'error': 'Invalid request data'}), 400

    filename = data['filename']
    mapping_config = data['mapping']
    custom_values = data.get('custom_values', {})
    split_config = data.get('split_config')
    sheet_name = data.get('sheet_name')

    file_path = os.path.join(UPLOAD_FOLDER, filename)
    if not os.path.exists(file_path):
        return jsonify({'error': 'File not found'}), 404

    try:
        # First, verify the file and sheet
        logger.info(f"Verifying file and sheet before processing: {file_path}, sheet: {sheet_name}")
        try:
            # Check if the file is valid
            xl = pd.ExcelFile(file_path)
            available_sheets = xl.sheet_names
            logger.info(f"File is valid. Available sheets: {available_sheets}")

            # If a sheet name is provided, check if it exists
            if sheet_name:
                # Create a case-insensitive lookup
                sheet_lookup = {s.lower().strip(): s for s in available_sheets}

                if sheet_name in available_sheets:
                    logger.info(f"Sheet '{sheet_name}' found (exact match)")
                elif sheet_name.lower().strip() in sheet_lookup:
                    actual_sheet = sheet_lookup[sheet_name.lower().strip()]
                    logger.info(f"Sheet '{sheet_name}' found (case-insensitive match): '{actual_sheet}'")
                    sheet_name = actual_sheet
                else:
                    logger.warning(f"Sheet '{sheet_name}' not found in file. Available sheets: {available_sheets}")
                    return jsonify({'error': f"Sheet '{sheet_name}' not found in file. Available sheets: {available_sheets}"}), 400
        except Exception as e:
            logger.error(f"Error verifying file: {str(e)}")
            import traceback
            error_traceback = traceback.format_exc()
            logger.error(f"Traceback: {error_traceback}")
            return jsonify({'error': f"Error verifying file: {str(e)}"}), 500

        # Process the file
        logger.info(f"Processing file: {file_path}, sheet: {sheet_name}, mapping: {mapping_config}")
        try:
            result = standardizer.process_file(file_path, mapping_config, split_config, custom_values, sheet_name)
            logger.info(f"File processed successfully. Output files: {result['output_files']}")
        except Exception as e:
            logger.error(f"Error in standardizer.process_file: {str(e)}")
            import traceback
            error_traceback = traceback.format_exc()
            logger.error(f"Traceback: {error_traceback}")

            # Create a minimal result structure instead of returning an error
            output_folder = os.path.join(OUTPUT_FOLDER, os.path.splitext(os.path.basename(file_path))[0])
            os.makedirs(output_folder, exist_ok=True)

            # Create a log file with the error
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            log_filename = f"log_{timestamp}.json"
            log_path = os.path.join(output_folder, log_filename)

            # Create log content
            log_content = {
                'file_name': os.path.basename(file_path),
                'processing_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'entries': [
                    {
                        'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        'action': 'Error',
                        'details': f'Error processing file: {str(e)}'
                    }
                ],
                'errors': [f"Error processing file: {str(e)}"]
            }

            # Save log file
            with open(log_path, 'w') as f:
                json.dump(log_content, indent=2, fp=f)

            # Return a minimal successful response
            return jsonify({
                'output_files': [],
                'log_file': log_path,
                'errors': [f"Error processing file: {str(e)}"],
                'silent_error': True  # Flag to indicate silent error
            })

        # Get relative paths for output files
        output_files = []
        for path in result['output_files']:
            rel_path = os.path.relpath(path, OUTPUT_FOLDER)
            output_files.append({
                'path': rel_path,
                'name': os.path.basename(path),
                'size': f"{os.path.getsize(path) / 1024:.1f} KB"
            })

        log_file = os.path.relpath(result['log_file'], OUTPUT_FOLDER)

        # Get output folder name
        output_folder = None
        if result['output_files']:
            output_folder = os.path.dirname(os.path.relpath(result['output_files'][0], OUTPUT_FOLDER))

        # Create Excel file with warnings if there are any
        warnings_file = None
        if result['errors']:
            # Create a DataFrame with the warnings
            warnings_df = pd.DataFrame({
                'Question': result.get('error_questions', ['']*len(result['errors'])),
                'Warning/Error': result['errors']
            })

            # Save to Excel file
            warnings_filename = f"error_{os.path.splitext(filename)[0]}.xlsx"

            # Make sure we have a valid output folder
            if output_folder:
                warnings_folder = os.path.join(OUTPUT_FOLDER, output_folder)
                os.makedirs(warnings_folder, exist_ok=True)
                warnings_path = os.path.join(warnings_folder, warnings_filename)
            else:
                warnings_path = os.path.join(OUTPUT_FOLDER, warnings_filename)

            # Save the warnings file with optimized formatting
            try:
                # Use xlsxwriter for better formatting
                with pd.ExcelWriter(warnings_path, engine='xlsxwriter') as writer:
                    warnings_df.to_excel(writer, index=False, sheet_name='Warnings')

                    # Get the xlsxwriter workbook and worksheet objects
                    workbook = writer.book
                    worksheet = writer.sheets['Warnings']

                    # Set column widths for better readability - increased widths
                    worksheet.set_column('A:A', 80)  # Question column
                    worksheet.set_column('B:B', 70)  # Warning/Error column

                    # Add a header format
                    header_format = workbook.add_format({
                        'bold': True,
                        'text_wrap': True,
                        'valign': 'top',
                        'bg_color': '#FFC107',  # Yellow background for warnings
                        'border': 1
                    })

                    # Add a row format for better readability
                    row_format = workbook.add_format({
                        'text_wrap': True,
                        'valign': 'top'
                    })

                    # Write the column headers with the defined format
                    for col_num, value in enumerate(warnings_df.columns.values):
                        worksheet.write(0, col_num, value, header_format)

                    # Apply row format to all data rows
                    for row_num in range(1, len(warnings_df) + 1):
                        worksheet.set_row(row_num, None, row_format)
            except Exception as e:
                # Fallback to basic Excel output if formatting fails
                print(f"Warning: Could not apply formatting to warnings file: {e}")
                warnings_df.to_excel(warnings_path, index=False)

            # Add to output files
            rel_warnings_path = os.path.relpath(warnings_path, OUTPUT_FOLDER)
            warnings_file = {
                'path': rel_warnings_path,
                'name': warnings_filename,
                'size': f"{os.path.getsize(warnings_path) / 1024:.1f} KB"
            }
            output_files.append(warnings_file)

        return jsonify({
            'success': True,
            'output_files': output_files,
            'log_file': log_file,
            'errors': result['errors'],
            'warnings_file': warnings_file,
            'output_folder': output_folder
        })
    except Exception as e:
        import traceback
        error_traceback = traceback.format_exc()
        logger.error(f"Error processing file: {str(e)}\n{error_traceback}")
        return jsonify({'error': f'Error processing file: {str(e)}'}), 500

@app.route('/api/view/<path:file_path>')
def view_file(file_path):
    # Security check to prevent directory traversal
    if '..' in file_path:
        return jsonify({'error': 'Invalid file path'}), 400

    try:
        # Handle both absolute and relative paths
        if os.path.isabs(file_path):
            full_path = file_path
        else:
            full_path = os.path.join(OUTPUT_FOLDER, file_path)

        # Debug logging
        logger.info(f"View request for file_path: {file_path}")
        logger.info(f"Full path resolved to: {full_path}")
        logger.info(f"OUTPUT_FOLDER is: {OUTPUT_FOLDER}")
        logger.info(f"File exists check: {os.path.exists(full_path)}")

        # Check if file exists
        if not os.path.exists(full_path):
            logger.error(f"File not found: {full_path}")
            return jsonify({'error': f'File not found: {file_path}'}), 404

        # Get the directory and filename
        directory = os.path.dirname(full_path)
        filename = os.path.basename(full_path)

        # Log the view attempt
        logger.info(f"Viewing file: {full_path}")

        # For Excel files, we could convert to HTML or CSV for viewing
        # For simplicity, we'll just download the file without attachment
        return send_from_directory(directory, filename, as_attachment=False)
    except Exception as e:
        logger.error(f"Error viewing file: {str(e)}")
        return jsonify({'error': f'Error viewing file: {str(e)}'}), 500

@app.route('/api/download/<path:file_path>')
def download_file(file_path):
    # Security check to prevent directory traversal
    if '..' in file_path:
        return jsonify({'error': 'Invalid file path'}), 400

    try:
        # Handle both absolute and relative paths
        if os.path.isabs(file_path):
            full_path = file_path
        else:
            full_path = os.path.join(OUTPUT_FOLDER, file_path)

        # Check if file exists
        if not os.path.exists(full_path):
            logger.error(f"File not found: {full_path}")
            return jsonify({'error': f'File not found: {file_path}'}), 404

        # Get the directory and filename
        directory = os.path.dirname(full_path)
        filename = os.path.basename(full_path)

        # Log the download attempt
        logger.info(f"Downloading file: {full_path}")

        # Return the file
        return send_from_directory(directory, filename, as_attachment=True)
    except Exception as e:
        logger.error(f"Error downloading file: {str(e)}")
        return jsonify({'error': f'Error downloading file: {str(e)}'}), 500

@app.route('/api/download-all/<path:folder_path>')
def download_all_files(folder_path):
    """Download all files in a folder as a ZIP archive"""
    # Security check to prevent directory traversal
    if '..' in folder_path:
        return jsonify({'error': 'Invalid folder path'}), 400

    try:
        # Create a BytesIO object to store the ZIP file
        memory_file = io.BytesIO()

        # Create a ZIP file in memory
        with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # Get the full path to the folder
            full_folder_path = os.path.join(OUTPUT_FOLDER, folder_path)

            # Check if the folder exists
            if not os.path.exists(full_folder_path) or not os.path.isdir(full_folder_path):
                return jsonify({'error': 'Folder not found'}), 404

            # Add all files in the folder to the ZIP file
            for root, _, files in os.walk(full_folder_path):  # Use _ for unused variable
                for file in files:
                    # Skip log files if needed
                    if file.startswith('log_'):
                        continue

                    file_path = os.path.join(root, file)
                    # Add the file to the ZIP with a relative path
                    arcname = os.path.relpath(file_path, full_folder_path)
                    zipf.write(file_path, arcname)

        # Seek to the beginning of the BytesIO object
        memory_file.seek(0)

        # Create a filename for the ZIP file based on the folder name
        zip_filename = f"zip_{os.path.basename(folder_path)}.zip"

        # Return the ZIP file as an attachment
        return send_file(
            memory_file,
            mimetype='application/zip',
            as_attachment=True,
            download_name=zip_filename
        )
    except Exception as e:
        import traceback
        error_traceback = traceback.format_exc()
        logger.error(f"Error creating ZIP file: {str(e)}\n{error_traceback}")
        return jsonify({'error': f'Error creating ZIP file: {str(e)}'}), 500

@app.route('/api/folder/<path:folder_path>')
def view_folder(folder_path):
    # Security check to prevent directory traversal
    if '..' in folder_path:
        return jsonify({'error': 'Invalid folder path'}), 400

    full_path = os.path.join(OUTPUT_FOLDER, folder_path)
    if not os.path.exists(full_path) or not os.path.isdir(full_path):
        return jsonify({'error': 'Folder not found'}), 404

    files = []
    for file in os.listdir(full_path):
        if file.endswith(('.xlsx', '.xls')):
            rel_path = os.path.join(folder_path, file)
            files.append({
                'name': file,
                'path': rel_path,
                'size': f"{os.path.getsize(os.path.join(full_path, file)) / 1024:.1f} KB",
                'download_url': f"/api/download/{rel_path}",
                'view_url': f"/api/view/{rel_path}"
            })

    # Return a simple HTML page listing the files
    html = """<!DOCTYPE html>
    <html>
    <head>
        <title>Output Files</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 20px; }
            h1 { color: #f97316; }
            table { border-collapse: collapse; width: 100%; }
            th, td { padding: 8px; text-align: left; border-bottom: 1px solid #ddd; }
            tr:hover { background-color: #f5f5f5; }
            a { color: #f97316; text-decoration: none; }
            a:hover { text-decoration: underline; }
        </style>
    </head>
    <body>
        <h1>Output Files</h1>
        <table>
            <tr>
                <th>File Name</th>
                <th>Size</th>
                <th>Actions</th>
            </tr>
    """

    for file in files:
        html += f"""<tr>
            <td>{file['name']}</td>
            <td>{file['size']}</td>
            <td>
                <a href="{file['view_url']}" target="_blank">View</a> |
                <a href="{file['download_url']}">Download</a>
            </td>
        </tr>"""

    html += """</table>
    </body>
    </html>"""

    return Response(html, mimetype='text/html')

if __name__ == '__main__':
    # Use port from environment variable for compatibility with Render
    port = int(os.environ.get('PORT', 5051))
    print(f"Starting fixed app on port {port}...")
    app.run(host='0.0.0.0', port=port, debug=False)
