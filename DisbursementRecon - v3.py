import pandas as pd
import numpy as np
from datetime import datetime
import re
import os
import tkinter as tk
from tkinter import filedialog
import logging

# Set up logging configuration
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Constants for column names
TRANSACTION_NARRATION = 'TRANSACTION NARRATION'
EFFECTIVE_DATE = 'EFFECTIVE DATE'
LOAN_NUMBER = 'LOAN NUMBER'
AMOUNT_DISBURSED = 'AMOUNT DISBURSED'
DESCRIPTION = 'Description'
DATE_COL = 'Date'
R_NUMBER = 'R-Number'
UNIQUE_REFERENCE = 'Unique Reference'

def select_file(message: str, filetypes=(("Excel files", "*.xlsx;*.xls"), ("All files", "*.*"))) -> str:
    """
    Display a file selection dialog with a custom message.

    Parameters:
        message (str): The message to display.
        filetypes (tuple): Tuple of file type filters.

    Returns:
        str: The selected file path.

    Raises:
        FileNotFoundError: If no file is selected.
    """
    logging.info(message)
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title=message, filetypes=filetypes)
    root.destroy()  # Clean up the Tkinter root window
    if not file_path:
        raise FileNotFoundError("No file was selected.")
    return file_path

def extract_r_number_updated(description: str) -> str:
    """
    Extract an R-number from a string using regex.

    Parameters:
        description (str): The text description to search.

    Returns:
        str: The extracted R-number or NaN if not found.
    """
    if isinstance(description, str):
        # Use re.IGNORECASE to handle lowercase 'r'
        r_number_match = re.search(r'(\d+R\d+)', description, re.IGNORECASE)
        return r_number_match.group() if r_number_match else np.nan
    return np.nan

def create_unique_reference(row: pd.Series) -> str:
    """
    Create a unique reference based on 'R-Number' and 'Amount'.

    Parameters:
        row (pd.Series): A row of data containing 'R-Number' and 'Amount'.

    Returns:
        str: The unique reference formatted as "<digits after R>-<Amount>" or NaN if fields are missing.
    """
    r_number = row.get(R_NUMBER)
    amount = row.get('Amount')  # Assuming bank statement has a column 'Amount'
    if pd.notna(r_number) and pd.notna(amount):
        # Ensure consistent formatting by converting to uppercase
        r_number = r_number.upper()
        if 'R' in r_number:
            # Extract digits after 'R'
            digits_after_R = r_number.split('R')[-1]
            # Format the absolute value of the amount to two decimals
            return f"{digits_after_R}-{abs(amount):.2f}"
        else:
            logging.warning(f"Unexpected R-number format: {r_number}")
    return np.nan

def process_disbursement_report() -> dict:
    """
    Process the disbursement report Excel file by filtering and calculating a unique reference.

    Returns:
        dict: A dictionary containing the processed DataFrame and the input file path.
    """
    disbursement_file_path = select_file("Upload the Disbursement Report")
    try:
        disbursement_df = pd.read_excel(disbursement_file_path, skiprows=6)
    except Exception as e:
        logging.error(f"Error reading disbursement report: {e}")
        raise

    # Check for necessary column
    if TRANSACTION_NARRATION not in disbursement_df.columns:
        logging.error(f"Column '{TRANSACTION_NARRATION}' not found in disbursement report.")
        raise KeyError(f"Column '{TRANSACTION_NARRATION}' not found in disbursement report.")

    # Filter out rows containing 'cash' or 'nan' in TRANSACTION NARRATION
    filtered_df = disbursement_df[
        ~disbursement_df[TRANSACTION_NARRATION].str.contains('cash|nan', case=False, na=False)
    ].copy()

    if EFFECTIVE_DATE not in filtered_df.columns:
        logging.error(f"Column '{EFFECTIVE_DATE}' not found in disbursement report.")
        raise KeyError(f"Column '{EFFECTIVE_DATE}' not found in disbursement report.")

    # Convert EFFECTIVE DATE column to datetime
    filtered_df[EFFECTIVE_DATE] = pd.to_datetime(filtered_df[EFFECTIVE_DATE], errors='coerce')

    # Remove rows where LOAN NUMBER or AMOUNT DISBURSED is NaN
    required_cols = [LOAN_NUMBER, AMOUNT_DISBURSED]
    filtered_df = filtered_df.dropna(subset=required_cols)

    # Create Unique Reference for disbursement data.
    # Converting LOAN NUMBER to int might fail if non-numeric characters exist.
    try:
        filtered_df[UNIQUE_REFERENCE] = (
            filtered_df[LOAN_NUMBER].astype(int).astype(str) + "-" +
            filtered_df[AMOUNT_DISBURSED].apply(lambda x: f"{round(x, 2):.2f}")
        )
    except Exception as e:
        logging.error(f"Error creating Unique Reference in disbursement report: {e}")
        raise

    return {'df': filtered_df, 'input_file_path': disbursement_file_path}

def process_bank_statement() -> dict:
    """
    Process the bank statement Excel file by filtering and calculating unique references.

    Returns:
        dict: A dictionary containing the processed DataFrame and the input file path.
    """
    input_file_path = select_file("Select the bank statement in the required format to upload")
    try:
        df = pd.read_excel(input_file_path)
    except Exception as e:
        logging.error(f"Error reading bank statement: {e}")
        raise

    if DESCRIPTION not in df.columns:
        logging.error(f"Column '{DESCRIPTION}' not found in bank statement.")
        raise KeyError(f"Column '{DESCRIPTION}' not found in bank statement.")

    # Filter out rows containing 'DEBIT TRANSFERST-' in Description
    df = df[~df[DESCRIPTION].str.contains('DEBIT TRANSFERST-', case=False, na=False)]

    if DATE_COL not in df.columns:
        logging.error(f"Column '{DATE_COL}' not found in bank statement.")
        raise KeyError(f"Column '{DATE_COL}' not found in bank statement.")

    # Convert Date column to datetime
    df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors='coerce')

    # Extract R-number and create Unique Reference
    df[R_NUMBER] = df[DESCRIPTION].apply(extract_r_number_updated)
    df[UNIQUE_REFERENCE] = df.apply(create_unique_reference, axis=1)

    return {'df': df, 'input_file_path': input_file_path}

def merge_dataframes(bank_df: pd.DataFrame, disbursement_df: pd.DataFrame):
    """
    Merge bank and disbursement DataFrames on Unique Reference and determine matched/unmatched rows.
    Dates are compared using their datetime values.

    Parameters:
        bank_df (pd.DataFrame): Processed bank statement data.
        disbursement_df (pd.DataFrame): Processed disbursement report data.

    Returns:
        tuple: A tuple containing (matched_data, unmatched_bank, unmatched_disbursement).
    """
    # Perform an outer join on Unique Reference
    all_data = pd.merge(bank_df, disbursement_df, on=UNIQUE_REFERENCE, how='outer', suffixes=('_bank', '_disbursement'))

    # Ensure date columns are datetime types
    all_data[DATE_COL] = pd.to_datetime(all_data[DATE_COL], errors='coerce')
    all_data[EFFECTIVE_DATE] = pd.to_datetime(all_data[EFFECTIVE_DATE], errors='coerce')

    # Calculate the absolute difference in days between the two dates
    all_data['date_diff'] = (all_data[DATE_COL] - all_data[EFFECTIVE_DATE]).abs().dt.days

    # Matched rows: where both dates exist and the difference is <= 7 days
    matched_data = all_data.dropna(subset=[DATE_COL, EFFECTIVE_DATE])
    matched_data = matched_data[matched_data['date_diff'] <= 7]

    # Unmatched rows:
    # - Rows with missing EFFECTIVE DATE are considered unmatched from the disbursement side.
    unmatched_bank = all_data[all_data[EFFECTIVE_DATE].isna()]
    # - Rows with missing Date are considered unmatched from the bank side.
    unmatched_disbursement = all_data[all_data[DATE_COL].isna()]

    # Note: Rows with both dates present but with a date difference >7 days are not included in matched_data.
    return matched_data, unmatched_bank, unmatched_disbursement

def main():
    """
    Main function to process the bank statement and disbursement report,
    merge data, and export unmatched data.
    """
    try:
        bank_data = process_bank_statement()
        bank_df = bank_data['df']

        disbursement_data = process_disbursement_report()
        disbursement_df = disbursement_data['df']

        matched_data, unmatched_bank, unmatched_disbursement = merge_dataframes(bank_df, disbursement_df)

        timestamp = datetime.now().strftime('%Y%m%d%H%M')
        output_directory = os.path.dirname(disbursement_data['input_file_path'])

        # Export unmatched bank data
        output_file_path_unmatched_bank = os.path.join(output_directory, f"Unmatched_Bank_{timestamp}.xlsx")
        unmatched_bank.to_excel(output_file_path_unmatched_bank, index=False)
        logging.info(f"Unmatched Bank data exported to: {output_file_path_unmatched_bank}")

        # Export unmatched disbursement data
        output_file_path_unmatched_disbursement = os.path.join(output_directory, f"Unmatched_Disbursement_{timestamp}.xlsx")
        unmatched_disbursement.to_excel(output_file_path_unmatched_disbursement, index=False)
        logging.info(f"Unmatched Disbursement data exported to: {output_file_path_unmatched_disbursement}")

        # Optionally log a preview of matched data
        logging.info("Matched Data (first 5 rows):")
        logging.info("\n" + str(matched_data.head()))

        logging.info("Processing completed successfully.")
    except Exception as e:
        logging.error(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
