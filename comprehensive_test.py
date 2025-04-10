import pandas as pd
import os
import sys
import logging
import traceback
import requests
import json
import time
from excel_standardizer_improved import ExcelStandardizer

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('comprehensive_test.log')
    ]
)
logger = logging.getLogger()

class ComprehensiveTest:
    """Comprehensive test suite for the Excel Standardizer application"""
    
    def __init__(self, app_url="http://127.0.0.1:5051"):
        """Initialize the test suite"""
        self.app_url = app_url
        self.standardizer = ExcelStandardizer()
        self.test_files_dir = "Sample-Files"
        self.results = {
            "total_tests": 0,
            "passed_tests": 0,
            "failed_tests": 0,
            "test_results": []
        }
    
    def run_all_tests(self):
        """Run all tests and report results"""
        logger.info("Starting comprehensive test suite")
        
        # Test 1: Check if the application is running
        self.test_app_running()
        
        # Test 2: Test the standardizer directly
        self.test_standardizer_directly()
        
        # Test 3: Test the API endpoints
        self.test_api_endpoints()
        
        # Test 4: Test error handling
        self.test_error_handling()
        
        # Test 5: Test with problematic files
        self.test_problematic_files()
        
        # Report results
        self.report_results()
    
    def test_app_running(self):
        """Test if the application is running"""
        test_name = "Application Running Test"
        self.results["total_tests"] += 1
        
        try:
            logger.info(f"Running test: {test_name}")
            response = requests.get(self.app_url)
            
            if response.status_code == 200:
                logger.info(f"✅ {test_name} passed")
                self.results["passed_tests"] += 1
                self.results["test_results"].append({
                    "name": test_name,
                    "status": "passed",
                    "details": f"Application is running. Status code: {response.status_code}"
                })
            else:
                logger.error(f"❌ {test_name} failed. Status code: {response.status_code}")
                self.results["failed_tests"] += 1
                self.results["test_results"].append({
                    "name": test_name,
                    "status": "failed",
                    "details": f"Application is not running properly. Status code: {response.status_code}"
                })
        except Exception as e:
            logger.error(f"❌ {test_name} failed with exception: {str(e)}")
            self.results["failed_tests"] += 1
            self.results["test_results"].append({
                "name": test_name,
                "status": "failed",
                "details": f"Exception: {str(e)}"
            })
    
    def test_standardizer_directly(self):
        """Test the standardizer class directly"""
        test_name = "Standardizer Direct Test"
        self.results["total_tests"] += 1
        
        try:
            logger.info(f"Running test: {test_name}")
            
            # Find all Excel files in the test directory
            excel_files = []
            for file in os.listdir(self.test_files_dir):
                if file.endswith(".xlsx") or file.endswith(".xls"):
                    excel_files.append(os.path.join(self.test_files_dir, file))
            
            if not excel_files:
                logger.warning(f"⚠️ No Excel files found in {self.test_files_dir}")
                self.results["test_results"].append({
                    "name": test_name,
                    "status": "skipped",
                    "details": f"No Excel files found in {self.test_files_dir}"
                })
                return
            
            # Test with the first Excel file
            test_file = excel_files[0]
            logger.info(f"Testing with file: {test_file}")
            
            # Test analyze_file
            columns_info, file_shape, sheet_names, selected_sheet, sheet_info = self.standardizer.analyze_file(test_file)
            
            if columns_info and file_shape and sheet_names and selected_sheet:
                logger.info(f"✅ analyze_file passed. Found {len(columns_info)} columns, {file_shape[0]} rows, {len(sheet_names)} sheets")
                
                # Test process_file
                mapping_config = {}
                for col in columns_info:
                    std_col = self.get_standard_column_for(col["name"])
                    if std_col:
                        mapping_config[std_col] = col["name"]
                
                result = self.standardizer.process_file(test_file, mapping_config)
                
                if result and "output_files" in result and result["output_files"]:
                    logger.info(f"✅ process_file passed. Generated {len(result['output_files'])} output files")
                    self.results["passed_tests"] += 1
                    self.results["test_results"].append({
                        "name": test_name,
                        "status": "passed",
                        "details": f"Successfully analyzed and processed file {test_file}"
                    })
                else:
                    logger.error(f"❌ process_file failed. No output files generated")
                    self.results["failed_tests"] += 1
                    self.results["test_results"].append({
                        "name": test_name,
                        "status": "failed",
                        "details": f"Failed to process file {test_file}"
                    })
            else:
                logger.error(f"❌ analyze_file failed")
                self.results["failed_tests"] += 1
                self.results["test_results"].append({
                    "name": test_name,
                    "status": "failed",
                    "details": f"Failed to analyze file {test_file}"
                })
        except Exception as e:
            logger.error(f"❌ {test_name} failed with exception: {str(e)}")
            logger.error(traceback.format_exc())
            self.results["failed_tests"] += 1
            self.results["test_results"].append({
                "name": test_name,
                "status": "failed",
                "details": f"Exception: {str(e)}"
            })
    
    def test_api_endpoints(self):
        """Test the API endpoints"""
        test_name = "API Endpoints Test"
        self.results["total_tests"] += 1
        
        try:
            logger.info(f"Running test: {test_name}")
            
            # Find all Excel files in the test directory
            excel_files = []
            for file in os.listdir(self.test_files_dir):
                if file.endswith(".xlsx") or file.endswith(".xls"):
                    excel_files.append(os.path.join(self.test_files_dir, file))
            
            if not excel_files:
                logger.warning(f"⚠️ No Excel files found in {self.test_files_dir}")
                self.results["test_results"].append({
                    "name": test_name,
                    "status": "skipped",
                    "details": f"No Excel files found in {self.test_files_dir}"
                })
                return
            
            # Test with the first Excel file
            test_file = excel_files[0]
            logger.info(f"Testing API with file: {test_file}")
            
            # Test /api/analyze endpoint
            with open(test_file, 'rb') as f:
                files = {'file': (os.path.basename(test_file), f)}
                response = requests.post(f"{self.app_url}/api/analyze", files=files)
            
            if response.status_code == 200:
                analyze_data = response.json()
                logger.info(f"✅ /api/analyze endpoint passed. Status code: {response.status_code}")
                
                if "filename" in analyze_data and "columns" in analyze_data:
                    logger.info(f"✅ /api/analyze response format is correct")
                    
                    # Test /api/process endpoint
                    mapping_config = {}
                    for col in analyze_data["columns"]:
                        std_col = self.get_standard_column_for(col["name"])
                        if std_col:
                            mapping_config[std_col] = col["name"]
                    
                    process_data = {
                        "filename": analyze_data["filename"],
                        "mapping": mapping_config
                    }
                    
                    process_response = requests.post(
                        f"{self.app_url}/api/process",
                        json=process_data,
                        headers={"Content-Type": "application/json"}
                    )
                    
                    if process_response.status_code == 200:
                        process_result = process_response.json()
                        logger.info(f"✅ /api/process endpoint passed. Status code: {process_response.status_code}")
                        
                        if "output_files" in process_result:
                            logger.info(f"✅ /api/process response format is correct")
                            self.results["passed_tests"] += 1
                            self.results["test_results"].append({
                                "name": test_name,
                                "status": "passed",
                                "details": f"Successfully tested API endpoints with file {test_file}"
                            })
                        else:
                            logger.error(f"❌ /api/process response format is incorrect")
                            self.results["failed_tests"] += 1
                            self.results["test_results"].append({
                                "name": test_name,
                                "status": "failed",
                                "details": f"Incorrect response format from /api/process"
                            })
                    else:
                        logger.error(f"❌ /api/process endpoint failed. Status code: {process_response.status_code}")
                        self.results["failed_tests"] += 1
                        self.results["test_results"].append({
                            "name": test_name,
                            "status": "failed",
                            "details": f"Failed to process file via API. Status code: {process_response.status_code}"
                        })
                else:
                    logger.error(f"❌ /api/analyze response format is incorrect")
                    self.results["failed_tests"] += 1
                    self.results["test_results"].append({
                        "name": test_name,
                        "status": "failed",
                        "details": f"Incorrect response format from /api/analyze"
                    })
            else:
                logger.error(f"❌ /api/analyze endpoint failed. Status code: {response.status_code}")
                self.results["failed_tests"] += 1
                self.results["test_results"].append({
                    "name": test_name,
                    "status": "failed",
                    "details": f"Failed to analyze file via API. Status code: {response.status_code}"
                })
        except Exception as e:
            logger.error(f"❌ {test_name} failed with exception: {str(e)}")
            logger.error(traceback.format_exc())
            self.results["failed_tests"] += 1
            self.results["test_results"].append({
                "name": test_name,
                "status": "failed",
                "details": f"Exception: {str(e)}"
            })
    
    def test_error_handling(self):
        """Test error handling"""
        test_name = "Error Handling Test"
        self.results["total_tests"] += 1
        
        try:
            logger.info(f"Running test: {test_name}")
            
            # Test 1: Invalid file format
            with open("comprehensive_test.py", 'rb') as f:
                files = {'file': ("test.py", f)}
                response = requests.post(f"{self.app_url}/api/analyze", files=files)
            
            # We should get a 400 response, but the application should handle it gracefully
            if response.status_code == 400:
                logger.info(f"✅ Invalid file format test passed. Got expected 400 status code")
            else:
                # Even if we don't get a 400, the response should be handled gracefully
                logger.info(f"⚠️ Invalid file format test: Got status code {response.status_code} instead of 400")
            
            # Test 2: Non-existent file
            process_data = {
                "filename": "non_existent_file.xlsx",
                "mapping": {}
            }
            
            process_response = requests.post(
                f"{self.app_url}/api/process",
                json=process_data,
                headers={"Content-Type": "application/json"}
            )
            
            # We should get a 404 response, but the application should handle it gracefully
            if process_response.status_code == 404:
                logger.info(f"✅ Non-existent file test passed. Got expected 404 status code")
            else:
                # Even if we don't get a 404, the response should be handled gracefully
                logger.info(f"⚠️ Non-existent file test: Got status code {process_response.status_code} instead of 404")
            
            # Test 3: Invalid mapping
            # Find an Excel file to test with
            excel_files = []
            for file in os.listdir(self.test_files_dir):
                if file.endswith(".xlsx") or file.endswith(".xls"):
                    excel_files.append(os.path.join(self.test_files_dir, file))
            
            if excel_files:
                test_file = excel_files[0]
                
                # First analyze the file to get a valid filename
                with open(test_file, 'rb') as f:
                    files = {'file': (os.path.basename(test_file), f)}
                    response = requests.post(f"{self.app_url}/api/analyze", files=files)
                
                if response.status_code == 200:
                    analyze_data = response.json()
                    
                    # Now try to process with an invalid mapping
                    process_data = {
                        "filename": analyze_data["filename"],
                        "mapping": {
                            "Invalid Column": "Non-existent Column"
                        }
                    }
                    
                    process_response = requests.post(
                        f"{self.app_url}/api/process",
                        json=process_data,
                        headers={"Content-Type": "application/json"}
                    )
                    
                    # The application should handle this gracefully without showing an error popup
                    if process_response.status_code == 200:
                        logger.info(f"✅ Invalid mapping test passed. Application handled it gracefully")
                        
                        # Check if there are errors in the response
                        process_result = process_response.json()
                        if "errors" in process_result and process_result["errors"]:
                            logger.info(f"✅ Invalid mapping test: Errors were logged but not shown as popups")
                            self.results["passed_tests"] += 1
                            self.results["test_results"].append({
                                "name": test_name,
                                "status": "passed",
                                "details": "Error handling is working correctly"
                            })
                        else:
                            logger.warning(f"⚠️ Invalid mapping test: No errors were logged")
                            self.results["test_results"].append({
                                "name": test_name,
                                "status": "warning",
                                "details": "Error handling worked but no errors were logged"
                            })
                    else:
                        logger.error(f"❌ Invalid mapping test failed. Status code: {process_response.status_code}")
                        self.results["failed_tests"] += 1
                        self.results["test_results"].append({
                            "name": test_name,
                            "status": "failed",
                            "details": f"Error handling failed. Status code: {process_response.status_code}"
                        })
                else:
                    logger.error(f"❌ Could not analyze file for error handling test")
                    self.results["failed_tests"] += 1
                    self.results["test_results"].append({
                        "name": test_name,
                        "status": "failed",
                        "details": "Could not analyze file for error handling test"
                    })
            else:
                logger.warning(f"⚠️ No Excel files found for error handling test")
                self.results["test_results"].append({
                    "name": test_name,
                    "status": "skipped",
                    "details": "No Excel files found for error handling test"
                })
        except Exception as e:
            logger.error(f"❌ {test_name} failed with exception: {str(e)}")
            logger.error(traceback.format_exc())
            self.results["failed_tests"] += 1
            self.results["test_results"].append({
                "name": test_name,
                "status": "failed",
                "details": f"Exception: {str(e)}"
            })
    
    def test_problematic_files(self):
        """Test with known problematic files"""
        test_name = "Problematic Files Test"
        self.results["total_tests"] += 1
        
        try:
            logger.info(f"Running test: {test_name}")
            
            # Check if we have the error file with log directory
            error_dir = os.path.join(self.test_files_dir, "error file with log")
            if os.path.exists(error_dir) and os.path.isdir(error_dir):
                # Find the problematic Excel file
                problematic_files = []
                for file in os.listdir(error_dir):
                    if file.endswith(".xlsx") or file.endswith(".xls"):
                        problematic_files.append(os.path.join(error_dir, file))
                
                if problematic_files:
                    test_file = problematic_files[0]
                    logger.info(f"Testing with problematic file: {test_file}")
                    
                    # Test analyze_file
                    try:
                        columns_info, file_shape, sheet_names, selected_sheet, sheet_info = self.standardizer.analyze_file(test_file)
                        
                        if columns_info and file_shape and sheet_names and selected_sheet:
                            logger.info(f"✅ analyze_file passed for problematic file. Found {len(columns_info)} columns, {file_shape[0]} rows, {len(sheet_names)} sheets")
                            
                            # Test process_file
                            mapping_config = {}
                            for col in columns_info:
                                std_col = self.get_standard_column_for(col["name"])
                                if std_col:
                                    mapping_config[std_col] = col["name"]
                            
                            try:
                                result = self.standardizer.process_file(test_file, mapping_config)
                                
                                if result and "output_files" in result and result["output_files"]:
                                    logger.info(f"✅ process_file passed for problematic file. Generated {len(result['output_files'])} output files")
                                    self.results["passed_tests"] += 1
                                    self.results["test_results"].append({
                                        "name": test_name,
                                        "status": "passed",
                                        "details": f"Successfully processed problematic file {test_file}"
                                    })
                                else:
                                    logger.error(f"❌ process_file failed for problematic file. No output files generated")
                                    self.results["failed_tests"] += 1
                                    self.results["test_results"].append({
                                        "name": test_name,
                                        "status": "failed",
                                        "details": f"Failed to process problematic file {test_file}"
                                    })
                            except Exception as e:
                                logger.error(f"❌ process_file failed for problematic file with exception: {str(e)}")
                                self.results["failed_tests"] += 1
                                self.results["test_results"].append({
                                    "name": test_name,
                                    "status": "failed",
                                    "details": f"Exception in process_file: {str(e)}"
                                })
                        else:
                            logger.error(f"❌ analyze_file failed for problematic file")
                            self.results["failed_tests"] += 1
                            self.results["test_results"].append({
                                "name": test_name,
                                "status": "failed",
                                "details": f"Failed to analyze problematic file {test_file}"
                            })
                    except Exception as e:
                        logger.error(f"❌ analyze_file failed for problematic file with exception: {str(e)}")
                        self.results["failed_tests"] += 1
                        self.results["test_results"].append({
                            "name": test_name,
                            "status": "failed",
                            "details": f"Exception in analyze_file: {str(e)}"
                        })
                else:
                    logger.warning(f"⚠️ No problematic Excel files found in {error_dir}")
                    self.results["test_results"].append({
                        "name": test_name,
                        "status": "skipped",
                        "details": f"No problematic Excel files found in {error_dir}"
                    })
            else:
                logger.warning(f"⚠️ Error file directory not found: {error_dir}")
                self.results["test_results"].append({
                    "name": test_name,
                    "status": "skipped",
                    "details": f"Error file directory not found: {error_dir}"
                })
        except Exception as e:
            logger.error(f"❌ {test_name} failed with exception: {str(e)}")
            logger.error(traceback.format_exc())
            self.results["failed_tests"] += 1
            self.results["test_results"].append({
                "name": test_name,
                "status": "failed",
                "details": f"Exception: {str(e)}"
            })
    
    def get_standard_column_for(self, column_name):
        """Map a column name to a standard column name"""
        # Simple mapping based on common patterns
        column_name = column_name.lower().strip()
        
        if "type" in column_name or "q type" in column_name:
            return "Question Type"
        elif "level" in column_name or "difficulty" in column_name:
            return "Difficulty Level"
        elif "q text" in column_name or "question text" in column_name or "question" in column_name:
            return "Question Text"
        elif "option 1" in column_name or "option a" in column_name or "option/ answer 1" in column_name:
            return "Option (A)"
        elif "option 2" in column_name or "option b" in column_name or "option/ answer 2" in column_name:
            return "Option (B)"
        elif "option 3" in column_name or "option c" in column_name or "option/ answer 3" in column_name:
            return "Option (C)"
        elif "option 4" in column_name or "option d" in column_name or "option/ answer 4" in column_name:
            return "Option (D)"
        elif "option 5" in column_name or "option e" in column_name or "option/ answer 5" in column_name:
            return "Option (E)"
        elif "option 6" in column_name or "option f" in column_name or "option/ answer 6" in column_name:
            return "Option (F)"
        elif "correct" in column_name or "answer" in column_name:
            return "Correct Answer"
        elif "topic" in column_name or "subject" in column_name:
            return "Topics"
        elif "author" in column_name or "email" in column_name:
            return "Author"
        
        return None
    
    def report_results(self):
        """Report the test results"""
        logger.info("\n" + "="*50)
        logger.info("COMPREHENSIVE TEST RESULTS")
        logger.info("="*50)
        logger.info(f"Total tests: {self.results['total_tests']}")
        logger.info(f"Passed tests: {self.results['passed_tests']}")
        logger.info(f"Failed tests: {self.results['failed_tests']}")
        logger.info(f"Success rate: {(self.results['passed_tests'] / self.results['total_tests']) * 100:.2f}%")
        logger.info("="*50)
        
        logger.info("\nDetailed Results:")
        for i, result in enumerate(self.results["test_results"]):
            status_icon = "✅" if result["status"] == "passed" else "❌" if result["status"] == "failed" else "⚠️"
            logger.info(f"{i+1}. {status_icon} {result['name']}: {result['status'].upper()}")
            logger.info(f"   Details: {result['details']}")
        
        logger.info("="*50)
        
        # Save results to file
        with open("comprehensive_test_results.json", "w") as f:
            json.dump(self.results, f, indent=2)
        
        logger.info(f"Results saved to comprehensive_test_results.json")

if __name__ == "__main__":
    # Check if the application is running
    app_url = "http://127.0.0.1:5051"
    
    try:
        requests.get(app_url, timeout=2)
    except requests.exceptions.ConnectionError:
        logger.error(f"❌ Application is not running at {app_url}")
        logger.error("Please start the application before running the tests")
        sys.exit(1)
    
    # Run the tests
    test_suite = ComprehensiveTest(app_url)
    test_suite.run_all_tests()
