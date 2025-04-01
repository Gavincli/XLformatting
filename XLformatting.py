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

# formats each column with bold font and proper size
for col in ['A', 'B', 'C', 'D', 'F', 'G']:
    sheet[f'{col}1'].font = Font(bold = True)
    sheet.column_dimensions[col].width = len(sheet[f'{col}1'].value) + 5

# Saves the new file
Poorly_Organized_Data_1.save('formatted_grades.xlsx')
