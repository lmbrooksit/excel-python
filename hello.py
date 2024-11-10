
from openpyxl.workbook import Workbook # type: ignore
from openpyxl import load_workbook # type: ignore

#Create a workbook object
# wb = Workbook()



# load existing spreadsheet

wb = load_workbook('hello.xlsx')

ws = wb.active

print("hello World!")