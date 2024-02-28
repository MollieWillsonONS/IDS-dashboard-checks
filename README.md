Metadata IDS Spreadsheet Checker

This Python script is designed to perform a series of checks on Excel files within a specified directory to ensure data consistency and adherence to naming standards. It logs the results of each check into a master log CSV file and moves files to separate directories based on the outcome of the checks (successful or failed).

Author
- Mollie Willson
- Email: mollie.willson@ons.gov.uk
- Date: January 2024

Dependencies
- pandas: For data manipulation and analysis.
- datetime: For generating timestamps and formatting dates.
- os: For operating system-related functions such as file manipulation.
- re: For regular expression pattern matching.
- glob: For searching directories using wildcard patterns.
- shutil: For file operations like moving files between directories.

Usage
- Ensure that the necessary libraries are installed.
- Set the directory paths for Excel files, successful submissions, and failed submissions in the script.
- Run the script to perform the checks and log results.

Notes
- Review the code comments for detailed explanations of each check and functionality.
- Customize the directory paths according to your specific file organization.
- Ensure that Excel files in the specified directory follow the expected format for accurate checks and results logging.