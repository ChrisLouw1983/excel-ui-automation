# Excel UI Automation

A utility for reconciling bank statements with disbursement reports. The tool processes Excel files and outputs any unmatched records.

## Setup

```bash
pip install -r requirements.txt
```

## Usage

Run the tool from the command line or launch the graphical interface.
If file paths are not provided in CLI mode, a file selection dialog will appear.

```bash
# CLI
python reconciliation_tool.py --bank path/to/bank.xlsx --disbursement path/to/report.xlsx --output output_directory

# GUI
 
python reconciliation_tool.py --gui
```

In the GUI, browse for the Bank Statement and Disbursement Report, then click
**Reconcile Now**. The unmatched records are saved in the current working
directory. A confirmation message shows the file paths and closes the window
after you acknowledge it.

