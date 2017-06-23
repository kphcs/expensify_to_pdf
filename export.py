from __future__ import print_function
from os.path import join, dirname, abspath
import xlrd
from datetime import datetime
from urllib import request
import math
from fpdf import FPDF
import os


declaratienummer = ''
tijdstip = ''
bedrag = ''
valuta = ''
merchant =''
extention =''


fname = join(dirname(dirname(abspath(__file__))), 'declaraties-expensify/', 'Bulk_Expense_Export-3.xls')

# Open the workbook
xl_workbook = xlrd.open_workbook(fname)

# List sheet names, and pull a sheet by name
#
sheet_names = xl_workbook.sheet_names()
print('Sheet Names', sheet_names)

xl_sheet = xl_workbook.sheet_by_name(sheet_names[0])

# Or grab the first sheet by index
#  (sheets are zero-indexed)
#
xl_sheet = xl_workbook.sheet_by_index(0)
print ('Sheet name: %s' % xl_sheet.name)

# Pull the first row by index
#  (rows/columns are also zero-indexed)
#
row = xl_sheet.row(0)  # 1st row

# Print 1st row values and types
#
from xlrd.sheet import ctype_text

print('(Column #) type:value')
for idx, cell_obj in enumerate(row):
    cell_type_str = ctype_text.get(cell_obj.ctype, 'unknown type')
    print('(%s) %s %s' % (idx, cell_type_str, cell_obj.value))

# Print all values, iterating through rows and columns
#
num_cols = xl_sheet.ncols   # Number of columns
for row_idx in range(0, xl_sheet.nrows):    # Iterate through rows
    print ('-'*40)
    print ('Row: %s' % row_idx)   # Print row number
    for col_idx in range(0, num_cols):  # Iterate through columns

        cell_obj = xl_sheet.cell(row_idx, col_idx)  # Get cell object by row, col
        print ('Column: [%s] cell_obj: [%s]' % (col_idx, cell_obj))

        if row_idx != 0:
            if col_idx == 0: declaratienummer = int(cell_obj.value)
            if col_idx == 1: tijdstip = xlrd.xldate.xldate_as_datetime(cell_obj.value, 0).strftime("%Y-%m-%d")
            if col_idx == 2: merchant = cell_obj.value
            if col_idx == 10: bedrag = cell_obj.value
            if col_idx == 9: valuta = cell_obj.value
            if col_idx == 11:

                if "pdf" in str(cell_obj.value):
                    extention = ".pdf"
                else:
                    extention = ".jpg"

                print ('Downloading %s' % cell_obj.value )
                filename = str(declaratienummer) + ' declaratie KP ' + str(merchant) + ' ' + str(tijdstip) + ' ' + str(bedrag) + ' ' + str(valuta) + str(extention)
                print ('%s' % filename)
                request.urlretrieve(cell_obj.value, filename)

                if "pdf" not in str(cell_obj.value):

                    pdf = FPDF(orientation = 'P', unit = 'mm', format='A4')
                    # compression is not yet supported in py3k version
                    pdf.compress = False
                    pdf.add_page()
                    # Unicode is not yet supported in the py3k version; use windows-1252 standard font
                    pdf.set_font('Arial', '', 14)
                    pdf.ln(10)
                    #pdf.write(5, filename)
                    pdf.image(filename, 0,0, 210, 297 )
                    pdf.output(filename + ".pdf", "F")
                    os.remove(filename)
