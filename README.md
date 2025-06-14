# Excel UI Automation

A utility for reconciling bank statements with disbursement reports. The tool processes Excel files and outputs any unmatched records.

## Setup

```bash
pip install -r requirements.txt
```

## Usage

Run the tool from the command line. If file paths are not provided, a file selection dialog will appear.

```bash
python reconciliation_tool.py --bank path/to/bank.xlsx --disbursement path/to/report.xlsx --output output_directory
```
