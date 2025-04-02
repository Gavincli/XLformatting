# Mary Catherine Shepherd, Sam Jenson, Gavin Clifton, Ben Funk, Thomas Apke
# professor anderson, section 003


import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font

myWorkbook = openpyxl.load_workbook("excel file name")

currentSheet = myWorkbook.active

rowNum = 2

for row in currentSheet.iter_rows(min_row=rowNum, values_only=True):
    classTitle = row[0]
    if classTitle and classTitle not in myWorkbook.sheetnames:
        worksheet = myWorkbook.create_sheet(title=classTitle)
myWorkbook.save("updated exel file name")


#apply filters for each worksheet
for worksheet in myWorkbook.worksheets:
    max_row = worksheet.max_row
    worksheet.auto_filter.ref = f"A1:D{max_row}"

# Adding summary statistics for each class
for worksheet in myWorkbook.worksheets :
    grades = [row[3] for row in worksheet.iter_rows(min_row=2, values_only=True) if row[3] is not None]
    if grades:
        summary = [
            ("Highest Grade", max(grades)),
            ("Lowest Grade", min(grades)),
            ("Mean Grade", statistics.mean(grades)),
            ("Median Grade", statistics.median(grades)),
            ("Number of Students", len(grades)),
        ]
        for i, (title, value) in enumerate(summary, start=1):
            worksheet[f"F{i}"] = title
            worksheet[f"G{i}"] = value


# formats each column with bold font and proper size
for col in ['A', 'B', 'C', 'D', 'F', 'G']:
    sheet[f'{col}1'].font = Font(bold = True)
    sheet.column_dimensions[col].width = len(sheet[f'{col}1'].value) + 5

# Saves the new file
Poorly_Organized_Data_1.save('formatted_grades.xlsx')
