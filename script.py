# Read the xls file
# Parse the content
# Find anything with a ! in it
# return file name or count

# import pandas as pd
from openpyxl import load_workbook
import win32com.client as win32

import os
import re


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
        # print(file_path)
        try:
          convert_xls_to_xlsx(file_path)
        except:
          print('Error converting file: ' + file_path)

def read_file(file_name):
  # convert_xls_to_xlsx(file_name)
  wb = load_workbook(file_name)
  print(file_name + '\n')

  for sheet in wb.sheetnames:
    ws = wb[sheet]

    if '<Worksheet' not in str(ws):
      continue

    # Starts with = means formula
    # regex would be like [A-Z]{0-9}* : [A-Z]{0-9}* 
    # 


    for row in ws.iter_rows():
      for cell in row:
        cell_value = cell.value
        cross_sheet_pattern = r'^=[a-zA-Z \d]*![A-Z]\d*'
        formula_pattern = r'^=[A-Z]+\([A-Z]+\d+:[A-Z]+\d+\)'
        if not isinstance(cell_value, str):
          continue

        if re.match(cross_sheet_pattern, cell_value):
          print(f"Cross-Sheet-Reference; Sheet: {sheet}, Cell {cell.coordinate}: {cell_value}")
        elif re.match(formula_pattern, cell_value):
          print(f"Formula; Sheet: {sheet}, Cell {cell.coordinate}: {cell_value}")

  print('\n')

def read_files_in_directory(dir_path):
  for root, dirs, files in os.walk(dir_path):
    for file in files:
      file_path = os.path.join(root, file)

      if (file_path.endswith('.xlsx')):
        read_file(file_path)

  # ws = wb['budget']
  # print(ws['J46'].value)


  

# convert_all_spreadsheets_to_xlsx('C:\\Users\\anike\\OneDrive\\Documents\\UCSD\\ERSP\\Script\\spreadsheets')
# convert_xls_to_xlsx('C:\\Users\\anike\\OneDrive\\Documents\\UCSD\\ERSP\\Script\\cross_sheet.xls')
# convert_xls_to_xlsx('C:\\Users\\anike\\OneDrive\\Documents\\UCSD\\ERSP\\Script\\no_cross_sheet.xls')
# read_file('C:\\Users\\anike\\OneDrive\\Documents\\UCSD\\ERSP\\Script\\cross_sheet.xls')
# read_file('C:\\Users\\anike\\OneDrive\\Documents\\UCSD\\ERSP\\Script\\spreadsheets\\database\\original\\0XLSBudgetingWGP2.xlsx')
read_files_in_directory('C:\\Users\\anike\\OneDrive\\Documents\\UCSD\\ERSP\\Script\\spreadsheets\\database')