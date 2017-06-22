# Import libraries and functions
import csv
import openpyxl
import datetime as dt
from openpyxl import Workbook
from openpyxl.compat import range
from openpyxl.utils import get_column_letter
from os.path import split, join

##############################################
# Open the file to format.
filename = 'C:\\Users\\kjeffris\\My Documents\\Excel\\export.csv'
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
dest_filename = join(new, newName)

# Create first sheet
ws1 = wb.worksheets[0]
ws1.title = "Raw Data"

# Copy contents of original output to new sheet
for row_index, row in enumerate(reader):
    for column_index, cell in enumerate(row):
        column_letter = get_column_letter((column_index + 1))
        ws1['%s%s' % (column_letter, (row_index + 1))].value = cell

ws2 = wb.create_sheet(title='Summary1')
ws2['A1'] = 'Analyte 1'
headerList = ['Sample', 'GNR1', 'Background', 'GNR1RFU', 'GNR2RFU', 'GNR3RFU',
              'RFU', 'RFUPercentCV', 'GNR1Signal',	'GNR2Signal', 'GNR3Signal',
              'Signal', 'Gnr1CalculatedConcentration',
              'Gnr2CalculatedConcentration',
              'Gnr3CalculatedConcentration', 'CalculatedConcentration',
              'CalculatedConcentrationPercentCV']
count = 0
for row in ws2.iter_rows('A{}:Q{}'.format(2, 2)):
    for cell in row:
        cell.value = headerList[count]
        count = count + 1

wb.save(filename=dest_filename)
