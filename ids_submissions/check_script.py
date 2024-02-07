# -*- coding: utf-8 -*-
"""


Created on Wed Jan 10 11:15:59 2024

@author: willsm
"""

import warnings; warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

import pandas as pd
from datetime import datetime
import os
import re
import glob
import shutil

# Defining the directory where Excel files are located
directory = r'C:\Users\willsm\scripts\architectproj\IDS-dashboard-checks\ids_submissions'

# Defining the directory where successful and failed submissions will be moved
success_directory = r'C:\Users\willsm\scripts\architectproj\IDS-dashboard-checks\ids_submissions\sucessful_submissions'
failure_directory = r'C:\Users\willsm\scripts\architectproj\IDS-dashboard-checks\ids_submissions\failed_submissions'

# Use glob to find all Excel files in the directory
xlsx_files = glob.glob(os.path.join(directory, '*.xlsx'))

# Function to generate a unique reference
def generate_unique_reference():
    return f"REF-{datetime.now().strftime('%Y%m%d%H%M%S')}"

# Function to log results to a master log CSV file
def log_results_to_csv(reference, date_submitted, data_creator, dataset_resource_name, gcp_dataset_name, business_catalogue_identifier, check_number, check_name, result, error_description=None, error_rows=None):
    log_df = pd.DataFrame({
        'Unique Reference': [reference],
        'Date Submitted': [date_submitted],
        'Dataset Resource Name': [dataset_resource_name],
        'GCP Dataset Name': [gcp_dataset_name],
        'IDS Business Catalogue Identifier': [business_catalogue_identifier],
        'Data Creator': [data_creator],
        'Check Number': [check_number],
        'Check Name': [check_name],
        'Check Result': [result],
        'Error Description': [error_description],
        'Error Rows': [error_rows]
    })

    # Check if the master log file exists, if not create it, else append to it
    if not os.path.isfile('master_log.csv'):
        log_df.to_csv('master_log.csv', index=False)
    else:
        log_df.to_csv('master_log.csv', mode='a', header=False, index=False)

# Iterate over each Excel file found
for file_path in xlsx_files:
    # Read the Excel file
    df = pd.read_excel(file_path, sheet_name=None)
    dataset_series = pd.read_excel(file_path, sheet_name='Dataset Series', skiprows=1)

    # Creating dataframes for the rest of the sheets
    dataset_resource = df['Dataset Resource']
    dataset_file = df['Dataset File']
    variables = df['Variables']
    codes_and_values = df['Codes and Values']

    # Extract relevant information from the 'Dataset Resource' sheet
    data_creator = df['Dataset Resource'].at[6, 'Value'].strip()
    dataset_resource_name = df['Dataset Resource'].at[0, 'Value'].strip()
    gcp_dataset_name = df['Dataset Resource'].at[44, 'Value'].strip()
    business_catalogue_identifier = df['Dataset Resource'].at[19, 'Value'].strip()

    
    # CHECK 1 - GCP BigQuery Dataset name must match across the dataset resource and dataset series tab. 

    error_rows_check1 = None
    # identifying cell that we want to check is the same as in the dataset resource column in the dataset_series dataframe
    gcp_name_cell = df['Dataset Resource'].at[44, 'Value']
    match_all_rows = all(dataset_series['Dataset Resource'].iloc[7:] == gcp_name_cell)

    # running code to see if the contents of all rows in dataset resource column match the contents of GCP name cell.
    # Using iloc to start looking from row 8 only as the previous 7 rows were informational 
    # printing output for if they do or do not match

    if match_all_rows:
        print("CHECK 1 PASSED: All rows in the 'Dataset Resource' column match the contents of the GCP name cell.")
    else:
        error_rows_check1 = dataset_series['Dataset Resource'].iloc[7:][dataset_series['Dataset Resource'].iloc[7:] != gcp_name_cell].index.tolist()
        print(f"CHECK 1 FAILED: Not all rows in the 'Dataset Resource' column match the contents of the GCP name cell. Error occurred in row(s): {error_rows_check1}")

    # CHECK 2 - Series names must match across dataset series, dataset file and variables tabs. 


    # Extract series names from each dataframe, excluding 'nan' values
    error_rows_check2 = None
    ds_series_name = dataset_series['Dataset Series Name'].iloc[5:].dropna().drop_duplicates()
    df_series_name = dataset_file['Dataset Series'].iloc[21:].dropna().drop_duplicates()
    v_series_name = variables['Dataset Series'].iloc[20:].dropna().drop_duplicates()

    # Define dictionaries to store series names
    series_names = {
        'Dataset Series': ds_series_name.tolist(),
        'Dataset File': df_series_name.tolist(),
        'Variables': v_series_name.tolist()
    }

    # Identify series names not present in all three tabs
    all_series_names = set(series_names['Dataset Series'] + series_names['Dataset File'] + series_names['Variables'])
    present_in_all_tabs = set(series_names['Dataset Series']) & set(series_names['Dataset File']) & set(series_names['Variables'])
    missing_in_some_tabs_check2 = all_series_names - present_in_all_tabs
    missing_tabs_check2 = {}

    for name in missing_in_some_tabs_check2:
        missing_tabs_check2[name] = [tab for tab, names in series_names.items() if name not in names]

    if not missing_in_some_tabs_check2:
        print("CHECK 2 PASSED: All series names are consistent across all tabs.")
    else:
        missing_info_check2 = ', '.join([f"{name} missing in {', '.join(tabs)}" for name, tabs in missing_tabs_check2.items()])
        print(f"CHECK 2 FAILED: Series names not consistent across all tabs. {missing_info_check2}")

    # CHECK 3 - File names must match across dataset file and variables tabs. 

    # Extracting the file names from each dataframe, excluding 'nan' values
    error_rows_check3 = None
    df_file_name = dataset_file['File path and name'].iloc[21:].dropna().drop_duplicates()
    v_file_name = variables['Dataset file name'].iloc[20:].dropna().drop_duplicates()

    # Define dictionaries to store file names
    file_names = {
        'Dataset File': df_file_name.tolist(),
        'Variables': v_file_name.tolist()
    }

    # Identify file names not present in both tabs
    all_file_names = set(df_file_name.tolist() + v_file_name.tolist())
    present_in_both_tabs = set(df_file_name) & set(v_file_name)
    missing_in_some_tabs_check3 = all_file_names - present_in_both_tabs
    missing_tabs_check3 = {}

    for name in missing_in_some_tabs_check3:
        missing_tabs_check3[name] = [tab for tab, names in file_names.items() if name not in names]

    if not missing_in_some_tabs_check3:
        print("CHECK 3 PASSED: All file names are consistent across Dataset File and Variables tabs.")
    else:
        missing_info_check3 = ', '.join([f"{name} missing in {', '.join(tabs)}" for name, tabs in missing_tabs_check3.items()])
        print(f"CHECK 3 FAILED: File names not consistent across Dataset File and Variables tabs. {missing_info_check3}")

    # CHECK 4 - All GCP names must conform to IDS GCP naming standards (no capital letters, no spaces, not starting with a number)

    # Extract GCP names from the 'Dataset Resource' column, excluding 'nan' values
    error_rows_check4 = None
    gcp_names = dataset_series['Dataset Resource'].iloc[7:].dropna()
    contains_capital_letters = any(char.isupper() for gcp_name in gcp_names for char in gcp_name)
    contains_spaces = any(char.isspace() for gcp_name in gcp_names for char in gcp_name)
    starts_with_number = any(char.isdigit() for gcp_name in gcp_names for char in gcp_name)
    leading_trailing_spaces = any(gcp_name.strip() != gcp_name for gcp_name in gcp_names)


    if not contains_capital_letters and not leading_trailing_spaces and not starts_with_number:
        print("CHECK 4 PASSED: All GCP names follow the naming standards.")
    else:
        error_rows_check4 = gcp_names[(gcp_names.str.contains(r'[A-Z]')) | (gcp_names.str.match(r'^\d'))].index.tolist()
        
        # Check for leading or trailing spaces and update error_rows_check4 accordingly
        error_rows_check4 += [index for index, gcp_name in gcp_names.items() if gcp_name.strip() != gcp_name]
        
        print(f"CHECK 4 FAILED: Some GCP names do not conform to the naming standards. Error occurred in row(s): {error_rows_check4}")

    reference = generate_unique_reference()
    date_submitted = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    # CHECK 5 - Verify that the lists in the search keywords cell are comma-separated and contain only letters, commas and spaces.
    error_rows_check5 = None
    cell_value = df['Dataset Resource'].at[16, 'Value']

    # Defining a regular expression pattern for valid content (letters, commas and spaces)
    valid_pattern = re.compile(r'^[a-zA-Z, ]+$')

    if valid_pattern.match(cell_value):
        print("CHECK 5 PASSED: The keywords list is formatted correctly.")
    else:
        error_rows_check5 = [16] 
        print(f"CHECK 5 FAILED: The keywords list contains invalid content. Error occurred in row(s): {error_rows_check5}")


    # Log results for each file
    reference = generate_unique_reference()
    date_submitted = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    # Log results for Check 1
    log_results_to_csv(
        reference,
        date_submitted,
        data_creator,
        dataset_resource_name,
        gcp_dataset_name,
        business_catalogue_identifier,
        1,
        "GCP BigQuery Dataset name consistency",
        "Pass" if match_all_rows else "Fail",
        "GCP name mismatch" if not match_all_rows else None,
        str(error_rows_check1) if not match_all_rows else None
    )

    # Log results for Check 2
    error_description_check2 = ', '.join([f"{name} missing in {', '.join(tabs)}" for name, tabs in missing_tabs_check2.items()]) if missing_in_some_tabs_check2 else None
    log_results_to_csv(
        reference,
        date_submitted,
        data_creator,
        dataset_resource_name,
        gcp_dataset_name,
        business_catalogue_identifier,
        2,
        "Series name consistency across tabs",
        "Pass" if not missing_in_some_tabs_check2 else "Fail",
        error_description_check2,
        str(missing_in_some_tabs_check2) if missing_in_some_tabs_check2 else None
    )

    # Log results for Check 3
    error_description_check3 = ', '.join([f"{name} missing in {', '.join(tabs)}" for name, tabs in missing_tabs_check3.items()]) if missing_in_some_tabs_check3 else None
    log_results_to_csv(
        reference,
        date_submitted,
        data_creator,
        dataset_resource_name,
        gcp_dataset_name,
        business_catalogue_identifier,
        3,
        "File name consistency across tabs",
        "Pass" if not missing_in_some_tabs_check3 else "Fail",
        error_description_check3,
        str(missing_in_some_tabs_check3) if missing_in_some_tabs_check3 else None
    )

    # Log results for Check 4
    error_description_check4 = "Some GCP names do not conform to the naming standards" if error_rows_check4 else None
    log_results_to_csv(
        reference,
        date_submitted,
        data_creator,
        dataset_resource_name,
        gcp_dataset_name,
        business_catalogue_identifier,
        4,
        "GCP Name conformity to IDS standards",
        "Pass" if not error_rows_check4 else "Fail",
        error_description_check4,
        str(error_rows_check4) if error_rows_check4 else None
    )

    # Log results for Check 5
    error_description_check5 = "The specified cell contains invalid content" if error_rows_check5 else None
    log_results_to_csv(
        reference,
        date_submitted,
        data_creator,
        dataset_resource_name,
        gcp_dataset_name,
        business_catalogue_identifier,
        5,  
        "Keywords list formatted correctly",
        "Pass" if not error_rows_check5 else "Fail",
        error_description_check5,
        str(error_rows_check5) if error_rows_check5 else None
    )

# Checking if all checks passed
    all_checks_passed = (
        match_all_rows and
        not missing_in_some_tabs_check2 and
        not missing_in_some_tabs_check3 and
        not error_rows_check4 and
        not error_rows_check5
    )

    if all_checks_passed:
        # Moves the file to the successful submissions directory
        shutil.move(file_path, success_directory)
    else:
        # Moves the file to the failed submissions directory
        shutil.move(file_path, failure_directory)