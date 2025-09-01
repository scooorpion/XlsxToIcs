# Excel to ICS Converter (For MUST students)

Convert Excel class schedules to Apple/Google calendar format (.ics)

由于目前WeMust仅支持导出一周课表的xlsx文件，因此本项目支持将多个xlsx文件整合为一个.ics文件。
该项目仅适用于iPhone与macOS用户，在macOS的Calender中导入.ics文件后，即可在iCloud中同步课表。

## How to Use

### Mac Users
1. Open Terminal
2. Navigate to the project folder: `cd /path/to/XlsxToIcs/ics`
3. Install required packages: `pip install -r requirements.txt`
4. Run: `python gui.py`
5. Click "Select Excel Files" to choose your files
6. Enable deduplication if needed
7. Click "Convert to ICS"
8. Your .ics file will be saved automatically



## Command Line Version (main.py)

For advanced users who prefer command line:

1. Edit `main.py` and modify the file path:
   ```python
   file_path = '/path/to/your/excel/file.xlsx'
   ```
2. Run: `python main.py`
3. The .ics file will be saved to your Downloads folder

**Note**: Command line version requires manual file path editing for each conversion.