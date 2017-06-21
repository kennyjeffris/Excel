# Import libraries and functions
from tkinter.filedialog import askopenfilename
import csv
import openpyxl
import datetime as dt
from openpyxl import Workbook
from openpyxl.compat import range
from openpyxl.utils import get_column_letter
from os.path import split, join

##############################################
# Open the file to format.
filename = askopenfilename()
f = open(filename)
# Change this when implementing into program
##############################################

# Set up CSV tools
csv.register_dialect('comma', delimiter=',')
reader = csv.reader(f, dialect='comma')

# Initialize .xlsx file
wb = Workbook()
(new, extra) = split(filename)
newName = 'output.xlsx'
print(new)
dest_filename = join(new, newName)

# Create first sheet
ws1 = wb.worksheets[0]
ws1.title = "Raw Data"

# Copy contents of original output to new sheet
for row_index, row in enumerate(reader):
    for column_index, cell in enumerate(row):
        column_letter = get_column_letter((column_index + 1))
        ws1['%s%s' % (column_letter, (row_index + 1))].value = cell

ws2 = wb.create_sheet(title='Concentration Data')
ws2['F5'] = 3.14

wb.save(filename=dest_filename)
