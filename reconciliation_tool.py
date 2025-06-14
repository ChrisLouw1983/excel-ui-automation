import argparse
import logging
import re
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd

# Logging configuration
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Constants
desc_col = 'Description'
date_col = 'Date'
r_number_col = 'R-Number'
unique_ref_col = 'Unique Reference'
loan_number_col = 'LOAN NUMBER'
amount_disbursed_col = 'AMOUNT DISBURSED'
transaction_narration_col = 'TRANSACTION NARRATION'
effective_date_col = 'EFFECTIVE DATE'
amount_col = 'Amount'

def extract_r_number(text: str) -> str:
    """Extracts an R-number from text."""
    if isinstance(text, str):
        match = re.search(r'(\d+R\d+)', text, re.IGNORECASE)
        if match:
            return match.group().upper()
    return np.nan

def create_unique_reference(row: pd.Series) -> str:
    """Construct unique reference from R-number and amount."""
    r_number = row.get(r_number_col)
    amount = row.get(amount_col)
    if pd.notna(r_number) and pd.notna(amount):
        digits = r_number.upper().split('R')[-1]
        return f"{digits}-{abs(amount):.2f}"
    return np.nan

def select_file(message: str) -> Path:
    """Prompt user for a file path."""
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title=message,
                                           filetypes=(("Excel files", "*.xlsx;*.xls"),
                                                      ("All files", "*.*")))
    root.destroy()
    if not file_path:
        raise FileNotFoundError("No file selected")
    return Path(file_path)

def process_disbursement_report(path: Path) -> pd.DataFrame:
    """Load and clean disbursement report."""
    df = pd.read_excel(path, skiprows=6)
    if transaction_narration_col not in df.columns:
        raise KeyError(f"Missing column {transaction_narration_col}")
    df = df[~df[transaction_narration_col].str.contains('cash|nan', case=False, na=False)].copy()
    df[effective_date_col] = pd.to_datetime(df[effective_date_col], errors='coerce')
    df = df.dropna(subset=[loan_number_col, amount_disbursed_col])
    df[unique_ref_col] = (
        df[loan_number_col].astype(int).astype(str) + '-' +
        df[amount_disbursed_col].apply(lambda x: f"{round(x, 2):.2f}")
    )
    return df

def process_bank_statement(path: Path) -> pd.DataFrame:
    """Load and clean bank statement."""
    df = pd.read_excel(path)
    if desc_col not in df.columns:
        raise KeyError(f"Missing column {desc_col}")
    df = df[~df[desc_col].str.contains('DEBIT TRANSFERST-', case=False, na=False)].copy()
    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
    df[r_number_col] = df[desc_col].apply(extract_r_number)
    df[unique_ref_col] = df.apply(create_unique_reference, axis=1)
    return df

def merge_frames(bank_df: pd.DataFrame, disb_df: pd.DataFrame):
    """Merge and split matched/unmatched records."""
    merged = pd.merge(bank_df, disb_df, on=unique_ref_col, how='outer',
                      suffixes=('_bank', '_disb'))
    merged[date_col] = pd.to_datetime(merged[date_col], errors='coerce')
    merged[effective_date_col] = pd.to_datetime(merged[effective_date_col], errors='coerce')
    merged['date_diff'] = (merged[date_col] - merged[effective_date_col]).abs().dt.days
    matched = merged.dropna(subset=[date_col, effective_date_col])
    matched = matched[matched['date_diff'] <= 7]
    unmatched_bank = merged[merged[effective_date_col].isna()]
    unmatched_disb = merged[merged[date_col].isna()]
    return matched, unmatched_bank, unmatched_disb

def reconcile(bank_path: Path, disb_path: Path, output_dir: Path):
    bank_df = process_bank_statement(bank_path)
    disb_df = process_disbursement_report(disb_path)
    matched, unmatched_bank, unmatched_disb = merge_frames(bank_df, disb_df)
    timestamp = datetime.now().strftime('%Y%m%d%H%M')
    unmatched_bank.to_excel(output_dir / f"Unmatched_Bank_{timestamp}.xlsx", index=False)
    unmatched_disb.to_excel(output_dir / f"Unmatched_Disbursement_{timestamp}.xlsx", index=False)
    logging.info("Reconciliation complete")
    logging.info("Matched sample:\n%s", matched.head())


def main():
    parser = argparse.ArgumentParser(description="Reconcile bank vs disbursement reports")
    parser.add_argument('--bank', type=Path, help='Path to bank statement Excel')
    parser.add_argument('--disbursement', type=Path, help='Path to disbursement report Excel')
    parser.add_argument('--output', type=Path, help='Directory for output files', default=Path.cwd())
    args = parser.parse_args()

    bank_path = args.bank or select_file("Select the bank statement")
    disb_path = args.disbursement or select_file("Select the disbursement report")

    reconcile(bank_path, disb_path, args.output)


if __name__ == '__main__':
    main()
