# Import libraries and functions
import csv
import string
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

##############################################
# Helper functions


def col2num(col):
    """Convert excel column letter to index number."""
    num = 0
    for c in col:
        if c in string.ascii_letters:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num


def getItems(sheet, analyte, feature):
    """Get the items in a column of a request sheet."""
    done = False
    count = 1
    while (not done):
        column_letter = get_column_letter(count)
        index = column_letter + '1'
        if feature == sheet[index].value:
            item = []
            for i in range(analyte, 65, 4):
                index = column_letter + str(i + 1)
                item.append(sheet[index].value)
            done = True
            return item
        count += 1
    return []


#############################################
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

########################################
# Copy Raw Data to sheet1
for row_index, row in enumerate(reader):
    for column_index, cell in enumerate(row):
        column_letter = get_column_letter((column_index + 1))
        ws1['%s%s' % (column_letter, (row_index + 1))].value = cell
########################################
analyteOrder = []
for index, row in enumerate(iterable=ws1.iter_rows(min_row=2, max_row=5,
                                                   max_col=1)):
    for cell in row:
        analyteOrder.append(cell.value)

headerList = ['Sample', 'Gnr1Background', 'Gnr1RFU', 'Gnr2RFU', 'Gnr3RFU',
              'RFU', 'RFUPercentCV', 'GNR1Signal',	'GNR2Signal', 'GNR3Signal',
              'Signal', 'Gnr1CalculatedConcentration',
              'Gnr2CalculatedConcentration',
              'Gnr3CalculatedConcentration', 'CalculatedConcentration',
              'CalculatedConcentrationPercentCV']

########################################
# Populate Summary1
ws2 = wb.create_sheet(title='Summary1')
for i in range(0, 4):
    analyteString = 'Analyte {} ({})'.format(i+1, analyteOrder[i])
    startRow = i*20+1
    ws2['A{}'.format(startRow)] = analyteString
    for index, col in enumerate(iterable=ws2.iter_cols(min_row=startRow+1,
                                min_col=1, max_row=startRow+1,
                                max_col=len(headerList))):
        for cell in col:
            cell.value = headerList[index]

    for index, row in enumerate(iterable=ws2.iter_rows(min_row=startRow+2,
                                max_col=1, max_row=startRow+17)):
        for cell in row:
            cell.value = index+1

    for index, col in enumerate(iterable=ws2.iter_cols(min_col=2,
                                                       max_col=len(headerList)-1)):
        feature = headerList[index + 1]
        values = getItems(ws1, i+1, feature)
        for index2, row in enumerate(iterable=ws2.iter_rows(min_col=index+2,
                                                            min_row=startRow+2,
                                                            max_row=startRow+17)):
            for cell in row:
                cell.value = values[index2]
########################################

wb.save(filename=dest_filename)
