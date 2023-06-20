# -*- coding: utf-8 -*-
import openpyxl
import xlwings
import pandas as pd

# ================================================FUNCTIONS================================================
# a function that convert number format into an excel formula
def apply_number_format(cell):
  if cell.value is not None and cell.number_format != 'General':
    cell.value = f'=IFERROR(TEXT("{cell.value}", "{cell.number_format}"), "{cell.value}")'
    cell.number_format = 'General'


# ==============================================MAIN PROGRAM===============================================
# path to the targeted file
path = '/path/to/file/example.xlsx'

# flatten existing equations
wb = openpyxl.load_workbook(path, data_only=True)

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
wb.save(path)

# open the excel file in the background with xlwings such that the formulas' value are computed by excel
excel_app = xlwings.App(visible=False)
excel_book = excel_app.books.open(path)
excel_book.save()
excel_book.close()
excel_app.quit()

# flatten all formulas
wb = openpyxl.load_workbook(path, data_only=True)
wb.save(path)

# store the excel file in pandas 
excel_file = pd.ExcelFile(path)

# Create an empty dictionary to store the DataFrames for each sheet
dfs = {}

# Iterate over each sheet name in the Excel file
for sheet_name in excel_file.sheet_names:
    # Read the sheet into a DataFrame with correct indexing and titles
    df = pd.read_excel(excel_file, sheet_name=sheet_name, index_col=None)
    # Store the DataFrame in the dictionary with the sheet name as the key
    dfs[sheet_name] = df

# integrate with other programs and get the result from dfs
print(dfs["DETAIL"])