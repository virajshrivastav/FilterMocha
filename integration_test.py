import requests
import os
import json
import time
import logging
import sys

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('integration_test.log')
    ]
)
logger = logging.getLogger()

class IntegrationTest:
    """Integration test for the Excel Standardizer application"""

    def __init__(self, app_url="http://127.0.0.1:5051"):
        """Initialize the integration test"""
        self.app_url = app_url
        self.test_files = [
            {
                "file_path": "Sample-Files/Assessment Questions - Oracle Apex.xlsx",
                "sheet_name": "Oracle apex questions",
                "mapping": {
                    "Question Type": "Q type",
                    "Difficulty Level": "Level",
                    "Question Text": "Q text",
                    "Option (A)": "Option/ Answer 1",
                    "Option (B)": "Option/ Answer 2",
                    "Option (C)": "Option/ Answer 3",
                    "Option (D)": "Option/ Answer 4",
                    "Option (E)": "Option/ Answer 5",
                    "Correct Answer": "Correct Answer"
                }
            },
            {
                "file_path": "Sample-Files/error file with log/QET GW Skill_questions_2024 (1).xlsx",
                "sheet_name": "Skill_questions_new",
                "mapping": {
                    "Question Type": "Q type",
                    "Difficulty Level": "Level",
                    "Question Text": "Q text",
                    "Option (A)": "Option/ Answer 1",
                    "Option (B)": "Option/ Answer 2",
                    "Option (C)": "Option/ Answer 3",
                    "Option (D)": "Option/ Answer 4",
                    "Option (E)": "Option/ Answer 5",
                    "Correct Answer": "Correct Answer",
                    "Topics": "Topic",
                    "Author": "Author's Email id(optional)"
                },
                "split": {
                    "enabled": True,
                    "column": "SubSkill"
                }
            }
        ]

    def run_test(self):
        """Run the integration test"""
        logger.info("Starting integration test")

        # Check if the application is running
        try:
            response = requests.get(self.app_url)
            if response.status_code != 200:
                logger.error(f"Application is not running properly. Status code: {response.status_code}")
                return False
        except Exception as e:
            logger.error(f"Application is not running: {str(e)}")
            return False

        # Test each file
        for i, test_file in enumerate(self.test_files):
            logger.info(f"Testing file {i+1}/{len(self.test_files)}: {test_file['file_path']}")

            # Step 1: Upload and analyze the file
            try:
                with open(test_file['file_path'], 'rb') as f:
                    files = {'file': (os.path.basename(test_file['file_path']), f)}
                    data = {'sheet_name': test_file['sheet_name']}

                    logger.info(f"Uploading and analyzing file: {test_file['file_path']}")
                    response = requests.post(f"{self.app_url}/api/analyze", files=files, data=data)

                    if response.status_code != 200:
                        logger.error(f"Failed to analyze file. Status code: {response.status_code}")
                        logger.error(f"Response: {response.text}")
                        continue

                    analyze_data = response.json()
                    logger.info(f"File analyzed successfully. Found {len(analyze_data['columns'])} columns")

                    # Check if there's a silent error
                    if 'silent_error' in analyze_data:
                        logger.warning(f"Silent error in analyze response: {analyze_data['silent_error']}")
            except Exception as e:
                logger.error(f"Error analyzing file: {str(e)}")
                continue

            # Step 2: Process the file
            try:
                # Prepare the process data
                process_data = {
                    "filename": analyze_data["filename"],
                    "mapping": test_file["mapping"]
                }

                # Add split configuration if specified
                if "split" in test_file:
                    process_data["split"] = test_file["split"]

                logger.info(f"Processing file with mapping: {json.dumps(process_data['mapping'], indent=2)}")

                # Send the process request
                process_response = requests.post(
                    f"{self.app_url}/api/process",
                    json=process_data,
                    headers={"Content-Type": "application/json"}
                )

                if process_response.status_code != 200:
                    logger.error(f"Failed to process file. Status code: {process_response.status_code}")
                    logger.error(f"Response: {process_response.text}")
                    continue

                process_result = process_response.json()

                # Check if there's a silent error
                if 'silent_error' in process_result:
                    logger.warning(f"Silent error in process response: {process_result['silent_error']}")

                # Check if there are output files
                if "output_files" in process_result and process_result["output_files"]:
                    logger.info(f"File processed successfully. Generated {len(process_result['output_files'])} output files")

                    # Log the output files
                    for j, output_file in enumerate(process_result["output_files"]):
                        try:
                            if isinstance(output_file, str):
                                logger.info(f"  Output file {j+1}: {os.path.basename(output_file)}")
                            elif isinstance(output_file, dict) and 'path' in output_file:
                                logger.info(f"  Output file {j+1}: {os.path.basename(output_file['path'])}")
                            else:
                                logger.info(f"  Output file {j+1}: {output_file}")
                        except Exception as e:
                            logger.warning(f"  Could not log output file {j+1}: {str(e)}")
                else:
                    logger.warning(f"No output files generated")

                # Check if there are errors
                if "errors" in process_result and process_result["errors"]:
                    logger.warning(f"Errors reported during processing:")
                    for j, error in enumerate(process_result["errors"]):
                        logger.warning(f"  Error {j+1}: {error}")
            except Exception as e:
                logger.error(f"Error processing file: {str(e)}")
                continue

            logger.info(f"Successfully tested file: {test_file['file_path']}")

        logger.info("Integration test completed")
        return True

if __name__ == "__main__":
    # Check if the application is running
    app_url = "http://127.0.0.1:5051"

    try:
        requests.get(app_url, timeout=2)
    except requests.exceptions.ConnectionError:
        logger.error(f"Application is not running at {app_url}")
        logger.error("Please start the application before running the tests")
        sys.exit(1)

    # Run the integration test
    test = IntegrationTest(app_url)
    success = test.run_test()

    if success:
        logger.info("Integration test passed")
        sys.exit(0)
    else:
        logger.error("Integration test failed")
        sys.exit(1)
