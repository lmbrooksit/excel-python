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

# Print something from the spreadsheet
print(ws["A2"])