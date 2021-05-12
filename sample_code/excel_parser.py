import sys,os
import pandas as pd
import numpy as np
import openpyxl
from openpyxl import load_workbook
import xlrd

path_info = "path to directory"
inputfile_name = "name of input file.csv"
outfile_name = "name of output file.xlsx"

df = pd.read_csv(os.path.join(path_info,inputfile_name))     ## this should have a column N with few cells blank and few cells with multiple values

writer = pd.ExcelWrite(os.path.join(path_info,outfile_name), engine = 'xlsxwriter')

#Convert the dataframe to an XlsxWriter Excel Object
df.to_excel(writer, sheet_name = 'sheet1', index = False)

#Get the xlsxwrite workbook and worksheet objects
workbook = writer.book
worksheet = writer.sheets['sheet1']

#Define the format for the row
cell_format =  workbook.add_format({'bg_color': 'red'})
cell_format2 = workbook.add_format({'bg_color': 'yellow'})

# Pick up the row index numbers where selected column has blank cellseg column N of df
'''
Use "|" or "&" to apply multiple criteria
'''
rows_with_blnk_cells = df.index[pd.isnull(df['N'])]   # | df['N'].astype(str).str.contains(","))  
print(rows_with_blnk_cells) 

rows_with_2vals = df.index[df['N'].astype(str).str.contains("/n"))]   # these cells have more than one vlaues with new line separation

wrap_format = workbook.add_format({'text_wrap': TRUE})

## Running a loop to highlight rows where the specified column has blank cells
for col in range(0,df.shape[1]):                                            ## iterate through every column of df
    for row in rows_with_blnk_cells:                                        
        if pd.isnull(df.iloc[row,col]):                                     ## if cell is blank you will get error, hence write None values
            worksheet.write(row+1, col, None, cell_format)
        else:
            worksheet.write(row+1, col, df.iloc[row,col],cell_format)

# For cells having multiple entries with same or different values, follow this
for col in range(0,df.shape[1]):
    for row in rows_with_2vals:
        if pd.isnull(df.iloc[row,col]):
            worksheet.write(row+1, col, None)
        elif df['N'].astype(str).str.split[0] != df['N'].astype(str).str.split[1]:
            worksheet.write(row+1, col, df.iloc[row,col], cell_format2)

writer.save()