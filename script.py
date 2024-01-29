# Read the xls file
# Parse the content
# Find anything with a ! in it
# return file name or count

# import pandas as pd
from openpyxl import load_workbook
import win32com.client as win32

import os


def convert_xls_to_xlsx(file_name):
  excel = win32.gencache.EnsureDispatch('Excel.Application')
  wb = excel.Workbooks.Open(file_name)

  wb.SaveAs(file_name + "x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
  wb.Close()                               #FileFormat = 56 is for .xls extension
  excel.Application.Quit()

def convert_all_spreadsheets_to_xlsx(dir_path):
  for root, dirs, files in os.walk(dir_path):
    for file in files:
      file_path = os.path.join(root, file)

      if (file_path.endswith('.xls') and not (os.path.isfile(file_path + 'x'))):
        print(file_path)
        convert_xls_to_xlsx(file_path)

def read_file(file_name):
  convert_xls_to_xlsx(file_name)
  wb = load_workbook(file_name + 'x')
  ws = wb['Sheet2']

  print(ws['A1'].value)

convert_all_spreadsheets_to_xlsx('C:\\Users\\anike\\OneDrive\\Documents\\UCSD\\ERSP\\Script\\spreadsheets\\database')
# convert_xls_to_xlsx('C:\\Users\\anike\\OneDrive\\Documents\\UCSD\\ERSP\\Script\\cross_sheet.xls')
# convert_xls_to_xlsx('C:\\Users\\anike\\OneDrive\\Documents\\UCSD\\ERSP\\Script\\no_cross_sheet.xls')
# read_file('C:\\Users\\anike\\OneDrive\\Documents\\UCSD\\ERSP\\Script\\cross_sheet.xls')