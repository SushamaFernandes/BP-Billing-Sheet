# AI Bill - Excel Processor

This is a Python desktop application using Tkinter that allows users to upload, process, and save Excel (.xlsx) files. It extracts specific columns, reformats dates, adds placeholder columns, and provides a simple GUI for file operations.

## Features
- Upload Excel (.xlsx) file via file dialog
- Extract columns: Entry Date, Resource Name, Task Name, Actul Work(hrs)
- Format Entry Date to dd-mm-yyyy
- Add placeholder columns: Issue#, Teams, Module, Task Type, Mandays, Billable
- Save processed file via Save As dialog
- Status label for success/error messages

## Requirements
- Python 3.x
- pandas
- tkinter

## How to Run
1. Install dependencies: `pip install pandas openpyxl`
2. Run the application: `python main.py`

## Notes
- The application uses tkinter for the GUI and pandas for Excel processing.
- Placeholder columns are added empty or with default values.
- The app can be closed after saving or used for repeated processing.
