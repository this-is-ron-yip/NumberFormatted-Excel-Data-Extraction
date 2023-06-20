# -*- coding: utf-8 -*-
import openpyxl
import subprocess
import pandas as pd

# ================================================FUNCTIONS================================================
# a function that convert number format into an excel formula
def apply_number_format(cell):
  if cell.value is not None and cell.number_format != 'General':
    cell.value = f'=IFERROR(TEXT("{cell.value}", "{cell.number_format}"), "{cell.value}")'
    cell.number_format = 'General'

# a function to trigger formula calculation and convert to xlsx type
def open_excel_with_libreoffice(file_name):
  output = subprocess.check_output(['libreoffice', '--convert-to','xlsx', '--outdir', '/', file_name])
  print(output.decode())

# remove the flattened file after data read by pandas
def remove_flattened_file(file_path):
  output = subprocess.check_output(['rm', file_path])
  print(output.decode())

# ==============================================MAIN PROGRAM===============================================
# directory and name of the target file
file_directory = '/path/to/file'
file_name = 'example.xlsx'

# full paths
file_path = file_directory + '/' + file_name
flattened_path = '/' + file_name

# flatten existing equations
wb = openpyxl.load_workbook(file_path, data_only=True)

# for every cell in the workbook, call the apply_number_format() function
for ws in wb:
  # iterate the merged cells
  for merged_ranges in ws.merged_cells:
    min_col, min_row, max_col, max_row = merged_ranges.bounds
    top_left_cell = ws.cell(row=min_row, column=min_col)
    top_left_cell = apply_number_format(top_left_cell)

  # iterate the unmerged cells
  for row in ws.iter_rows(min_row=ws.min_row, min_col=ws.min_column, max_row=ws.max_row, max_col=ws.max_column):
    for cell in row:
      cell = apply_number_format(cell)

# save the modified workbook
wb.save(file_path)

# trigger formula calculation with libreoffice
open_excel_with_libreoffice(file_name)

# flatten all formulas
wb = openpyxl.load_workbook(flattened_path, data_only=True)
wb.save(flattened_path)


# store the excel file in pandas 
excel_file = pd.ExcelFile(flattened_path)

# Create an empty dictionary to store the DataFrames for each sheet
dfs = {}

# Iterate over each sheet name in the Excel file
for sheet_name in excel_file.sheet_names:
    # Read the sheet into a DataFrame with correct indexing and titles
    df = pd.read_excel(excel_file, sheet_name=sheet_name, index_col=None)
    # Store the DataFrame in the dictionary with the sheet name as the key
    dfs[sheet_name] = df

# remove the extra flattened file
remove_flattened_file(flattened_path)

# integrate with other programs and get the result from dfs
print(dfs["DETAIL"])