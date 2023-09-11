import openpyxl
from openpyxl.styles import Font

# Load the workbook and select the active worksheet
wb = openpyxl.load_workbook('result.xlsx')
sheet = wb.active

# Iterate through each cell in the worksheet
for row in sheet.iter_rows():
    for cell in row:
        # Check if the cell contains a numeric value and if the value is greater than 0
        if isinstance(cell.value, (int, float)) and cell.value > 0:
            cell.font = Font(color="FF0000")  # Set the font color to red

# Save the changes
wb.save('result.xlsx')
