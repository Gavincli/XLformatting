# Mary Catherine Shepherd, Sam Jenson, Gavin Clifton, Ben Funk, Thomas Apke

pass

# formats each column with bold font and proper size
for col in ['A', 'B', 'C', 'D', 'F', 'G']:
    sheet[f'{col}1'].font = Font(bold = True)
    sheet.column_dimensions[col].width = len(sheet[f'{col}1'].value) + 5

# Saves the new file
Poorly_Organized_Data_1.save('formatted_grades.xlsx')