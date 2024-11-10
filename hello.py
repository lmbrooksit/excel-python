from openpyxl.workbook import Workbook # type: ignore
from openpyxl import load_workbook # type: ignore

#Create a workbook object
# wb = Workbook()


# load existing spreadsheet
# wb = load_workbook()

# load existing spreadsheet
wb = load_workbook('hello.xlsx')

# create an active worksheet
ws = wb.active

# Set a variable
name = ws["A3"].value
color = ws["B3"].value

# Print something from the spreadsheet
print(f'{name}: {color}')