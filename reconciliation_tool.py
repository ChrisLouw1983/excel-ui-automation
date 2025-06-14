import argparse
import logging
import re
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

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


class ReconciliationApp:
    """Simple Tkinter UI for running the reconciliation."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Reconciliation Tool")
        self.root.geometry("700x500")
        self.root.resizable(False, False)

        # Paths selected by the user
        self.bank_path = tk.StringVar()
        self.disb_path = tk.StringVar()

        # Configure logger
        self.logger = logging.getLogger("ReconciliationApp")
        self.logger.setLevel(logging.INFO)
        for h in self.logger.handlers[:]:
            self.logger.removeHandler(h)
        stream_handler = logging.StreamHandler()
        stream_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        self.logger.addHandler(stream_handler)

        self._create_widgets()

    def _create_widgets(self) -> None:
        """Create and place all UI widgets."""
        opts = {'padx': 10, 'pady': 10}

        instructions = (
            "Browse for your Bank Statement and Disbursement Report then click 'Reconcile Now'."
        )
        ttk.Label(self.root, text=instructions, wraplength=680, justify="left").grid(row=0, column=0, columnspan=3, sticky='w', **opts)

        ttk.Label(self.root, text="1. Bank Statement:").grid(row=1, column=0, sticky='e', **opts)
        ttk.Entry(self.root, textvariable=self.bank_path, width=60, state='readonly').grid(row=1, column=1, sticky='w', **opts)
        ttk.Button(self.root, text="Browse...", command=self._browse_bank).grid(row=1, column=2, sticky='w', **opts)

        ttk.Label(self.root, text="2. Disbursement Report:").grid(row=2, column=0, sticky='e', **opts)
        ttk.Entry(self.root, textvariable=self.disb_path, width=60, state='readonly').grid(row=2, column=1, sticky='w', **opts)
        ttk.Button(self.root, text="Browse...", command=self._browse_disb).grid(row=2, column=2, sticky='w', **opts)

        self.run_btn = ttk.Button(self.root, text="Reconcile Now", command=self._run, state='disabled')
        self.run_btn.grid(row=3, column=1, pady=20)

        self.status_var = tk.StringVar()
        ttk.Label(self.root, textvariable=self.status_var, foreground="blue", wraplength=680, justify="left").grid(row=5, column=0, columnspan=3, sticky='w', **opts)

        # Enable button when all paths set
        self.bank_path.trace_add('write', self._check_ready)
        self.disb_path.trace_add('write', self._check_ready)

    def _browse_bank(self) -> None:
        path = filedialog.askopenfilename(title="Select Bank Statement", filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        if path:
            self.bank_path.set(path)

    def _browse_disb(self) -> None:
        path = filedialog.askopenfilename(title="Select Disbursement Report", filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        if path:
            self.disb_path.set(path)

    def _check_ready(self, *args) -> None:
        if self.bank_path.get() and self.disb_path.get():
            self.run_btn.config(state='normal')
        else:
            self.run_btn.config(state='disabled')

    def _run(self) -> None:
        try:
            output_bank, output_disb = reconcile(
                Path(self.bank_path.get()),
                Path(self.disb_path.get()),
                Path.cwd(),
            )
            self.status_var.set('Reconciliation complete.')
            messagebox.showinfo(
                'Success',
                f'Reconciliation complete.\nFiles saved to:\n{output_bank}\n{output_disb}'
            )
            self.root.destroy()
        except Exception as exc:  # pragma: no cover - UI error handling
            self.logger.exception("Error during reconciliation")
            self.status_var.set(f'Error: {exc}')
            messagebox.showerror('Error', str(exc))

def reconcile(bank_path: Path, disb_path: Path, output_dir: Path) -> tuple[Path, Path]:
    """Process Excel files and write unmatched entries to disk."""
    bank_df = process_bank_statement(bank_path)
    disb_df = process_disbursement_report(disb_path)
    matched, unmatched_bank, unmatched_disb = merge_frames(bank_df, disb_df)
    timestamp = datetime.now().strftime('%Y%m%d%H%M')
    bank_out = output_dir / f"Unmatched_Bank_{timestamp}.xlsx"
    disb_out = output_dir / f"Unmatched_Disbursement_{timestamp}.xlsx"
    unmatched_bank.to_excel(bank_out, index=False)
    unmatched_disb.to_excel(disb_out, index=False)
    logging.info("Reconciliation complete")
    logging.info("Matched sample:\n%s", matched.head())
    return bank_out, disb_out


def main():
    parser = argparse.ArgumentParser(description="Reconcile bank vs disbursement reports")
    parser.add_argument('--bank', type=Path, help='Path to bank statement Excel')
    parser.add_argument('--disbursement', type=Path, help='Path to disbursement report Excel')
    parser.add_argument('--output', type=Path, help='Directory for output files', default=Path.cwd())
    parser.add_argument('--gui', action='store_true', help='Launch graphical interface')
    args = parser.parse_args()

    if args.gui:
        root = tk.Tk()
        app = ReconciliationApp(root)
        root.mainloop()
    else:
        bank = args.bank or select_file("Select the bank statement")
        disb = args.disbursement or select_file("Select the disbursement report")
        out_bank, out_disb = reconcile(bank, disb, args.output)
        print(f'Reconciliation complete. Files saved to:\n{out_bank}\n{out_disb}')


if __name__ == '__main__':
    main()
