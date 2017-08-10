from tkinter.filedialog import askopenfilename, asksaveasfilename
from tkinter import messagebox
import sys
import csv
import tkinter as tk
from openpyxl import Workbook
from openpyxl.compat import range
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, PatternFill, colors
from openpyxl.styles.borders import Border, Side

##############################
# Global variables
root = tk.Tk()
max_row = 0
max_col = 0
##############################

def main():
    file = get_file()
    wb = Workbook()
    wb = init_raw_data(wb)
    analytes = get_analytes(wb)
    wb = format_file(wb, analytes)
    save_file(file, wb)

def get_file():
    global root
    root.withdraw()
    root.iconbitmap('proteinsimple_logo_bt.ico')
    success = False
    while not success:
        try:
            filename = askopenfilename(title='Choose your data files',
                                       multiple=False, filetypes=(('CSV Files', '*.csv'), ('All Files', '*.*')))
            if not filename:
                sys.exit()
            elif not filename.endswith('.csv'):
                success = False
                messagebox.showerror(message="Invalid Filetype.",
                                     title="Failure")
            else:
                success = True
        except csv.Error as error:
            messagebox.showerror(message="Invalid Filetype.",
                                 title="Failure")

    file = open(filename)
    return file

def save_file(file, wb):
    newName = 'output.xlsx'

    dest_filename = asksaveasfilename(title='Save File.', filetypes=(('xlsx files', '*.xlsx'), ('all files', '*.*')),
                                      initialfile=newName)
    try:
        wb.save(filename=dest_filename)
    except PermissionError as e:
        messagebox.showerror(message="The file you are trying to overwrite is open. Close it and try again",
                             title="Failure")
        sys.exit()

    if not dest_filename:
        sys.exit()

def init_raw_data(wb, reader):
    # Create first sheet
    global max_col
    global max_row
    ws1 = wb.worksheets[0]
    ws1.title = 'Raw data'
    for row_index, row in enumerate(reader):
        for column_index, cell in enumerate(row):
            column_letter = get_column_letter((column_index + 1))
            ws1['%s%s' % (column_letter, (row_index + 1))].value = cell
            max_col = column_index + 1
        max_row = row_index + 1

def get_analytes(wb):
    try:
        analyte_order = []
        done = False
        count = 1
        while not done:
            letter = get_column_letter(count)
            ind = letter + '1'
            if wb.ws1[ind].value == "AnalyteName":
                '''for l in range(2, 6):
                    ind = letter + str(l)
                    try:
                        analyteOrder.append(str(wb.ws1[ind].value))
                    except Exception:
                        analyteOrder.append(wb.ws1[ind].value)
                break'''

                # REWRITE SMARTER CODE HERE FIXME

                if not analyte_order:
                    analyte_order.append(str(wb.ws1[ind].value))
                else:
                    pass #FIXME
            count += 1
            if count == max_col:
                raise ValueError('Item not found')
    except ValueError as error:
        messagebox.showerror(message="Missing Analyte Names.  Please export your data with "
                                     "this item included", title="Failure")
        sys.exit()
    return analyte_order

def format_file(wb, analytes):
    if len(analytes) == 1:
        from one_by_72 import format
    else:
        num_samples = get_num_samples(wb, analytes)
        if num_samples == 16:
            from four_by_16 import format
        elif num_samples == 36:
            from four_by_36 import format
    return format(wb, analytes)

if __name__ == '__main__':
    sys.exit(main())
