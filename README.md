# Excel to ICS Converter (For MUST students)

Convert Excel class schedules to Apple/Google calendar format (.ics)

## How to Use

### For Mac Users
1. Open Terminal
2. Navigate to the project folder: `cd /path/to/XlsxToIcs/ics`
3. Run: `python gui.py`
4. Click "Select Excel Files" to choose your files
5. Enable deduplication if needed
6. Click "Convert to ICS"
7. Your .ics file will be saved automatically

### For Windows Users
1. Open Command Prompt or PowerShell
2. Navigate to the project folder: `cd C:\path\to\XlsxToIcs\ics`
3. Run: `python gui.py`
4. Click "Select Excel Files" to choose your files
5. Enable deduplication if needed
6. Click "Convert to ICS"
7. Your .ics file will be saved automatically


## Command Line Version (main.py)

For advanced users who prefer command line:

1. Edit `main.py` and modify the file path:
   ```python
   file_path = '/path/to/your/excel/file.xlsx'
   ```
2. Run: `python main.py`
3. The .ics file will be saved to your Downloads folder

**Note**: Command line version requires manual file path editing for each conversion.