import pandas as pd
import os

def process_file(inputfile, outputpath):
  filename = inputfile.split(os.path.sep)[-1].split('.')[0].split('-')[0]
  df = pd.read_excel(inputfile, engine='openpyxl', skiprows=1)
  #df = pd.read_excel(inputfile, skiprows=1)

  to_rename = [x for x in df.columns if x.startswith('Saldo')]
  df.rename(columns={to_rename[0]: 'Saldos', 'DESCRIPCION': ''}, inplace=True)

  to_drop = ['No vencido', 'Riesgo Máx.:', 'NIVEL']
  df.drop(to_drop, axis=1, inplace=True)

  df = df[df['COD'].str.startswith('80') | df['COD'].str.startswith('00')].copy()

  order = ['', 'COD', 'Saldos', 'Vencido', '1 - 30 días', '31 - 60 días', '61 - 90 días', '91 - 180 días', '181 - 365 días', '1 - 2 años', '+ 2 años']

  df_order = df[order].copy()

  df_order.reset_index(drop=True, inplace=True)

  df_order['index'] = df_order.index

  vendors = [x for x in df_order['COD'].unique() if x.startswith('80')]
  df_vendors = df_order[df_order['COD'].isin(vendors)].copy()
  n_vendors = len(vendors)
  mapping = {}
  for i in range(n_vendors):
    name = df_vendors.iloc[i, df_order.columns.get_loc("")]
    start = df_vendors.iloc[i, df_order.columns.get_loc("index")] + 1
    if i < n_vendors-1 :
      end = df_vendors.iloc[i+1, df_order.columns.get_loc("index")]
    else:
      end = df_order.shape[0]
    if end > start:
      mapping[vendors[i]] = {'start': start, 'end': end, 'name': name}

  columns = ['Vencido', '1 - 30 días', '31 - 60 días', '61 - 90 días', '91 - 180 días', '181 - 365 días', '1 - 2 años',
             '+ 2 años']
  cols_to_elminate = ['91 - 180 días', '181 - 365 días', '1 - 2 años', '+ 2 años']

  active_vendors = list(mapping.keys())
  for ac_ve in active_vendors:
    vendor = df_order.iloc[mapping[ac_ve]['start']:mapping[ac_ve]['end']]
    mask = vendor['Vencido'] > 0
    vendor = vendor[mask].copy()
    vendor.sort_values(by=[''], inplace=True)
    new_line = {}
    for col in columns:
      new_line[col] = vendor[col].sum()
    to_drop = []
    for col in cols_to_elminate:
      if new_line[col] == 0:
        to_drop.append(col)
    line_end = pd.DataFrame(new_line, index=[1])
    df1 = pd.concat([vendor.iloc[:], line_end]).reset_index(drop=True)
    df1.drop(to_drop, axis=1, inplace=True)
    line = pd.DataFrame({"": mapping[ac_ve]['name']}, index=[1])
    df2 = pd.concat([line, df1.iloc[:]]).reset_index(drop=True)
    df2.drop('index', axis=1, inplace=True)
    xlsx_filename = filename + '_' + ac_ve + '.xlsx'
    outputname = os.path.join(outputpath, xlsx_filename)
    format_excel(df2, outputname)

from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
from openpyxl.styles import numbers
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

def auto_col_width(wrks):
  for col in wrks.columns:
     max_length = 0
     column = col[0].column # Get the column name
     # Since Openpyxl 2.6, the column name is  ".column_letter" as .column became the column number (1-based) 
     for cell in col:
         try: # Necessary to avoid error on empty cells
             if len(str(cell.value)) > max_length:
                 max_length = len(cell.value)
         except:
             pass
     adjusted_width = (max_length + 2) * 1.4
     wrks.column_dimensions[column].width = adjusted_width
     return wrks
 
def format_excel(data, name):
  wb = Workbook()
  ws = wb.active
  for r in dataframe_to_rows(data, index=False, header=True):
    ws.append(r)
  # Set auto width of columns
  for col in ws.columns:
    max_length = 0
    column = col[0].column # Get the column name
    # Since Openpyxl 2.6, the column name is  ".column_letter" as .column became the column number (1-based) 
    for cell in col:
      try:  # Necessary to avoid error on empty cells
        if len(str(cell.value)) > max_length:
          max_length = len(cell.value)
      except:
        pass
    adjusted_width = (max_length + 2) * 1.1
    ws.column_dimensions[get_column_letter(column)].width = adjusted_width
  # Put columns in grey
  row_count = ws.max_row
  greyFill = PatternFill(start_color='00C0C0C0',
                   end_color='00C0C0C0',
                   fill_type='solid')
  for col in range(5,8):
    for row in range(1, row_count):
      cell = ws.cell(row=row, column=col)
      cell.fill = greyFill
  # Center the columns B to end
  row_count = ws.max_row
  column_count = ws.max_column
  for col in range(2, column_count+1):
    for row in range(1, row_count+1):
      cell = ws.cell(row=row, column=col)
      cell.alignment = Alignment(horizontal="center", vertical="center")
  # Apply format numbers
  for col in range(3, column_count+1):
    for row in range(3, row_count+1):
      cell = ws.cell(row=row, column=col)
      cell.number_format = '#,###;-#,###;;@'
  # Apply format
  ws.insert_rows(1)
  ws['A1'] = 'SALDO CLIENTES POR ANTIGÜEDAD'
  row_count = ws.max_row
  column_count = ws.max_column

  ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=column_count)
  top_left_cell = ws['A1']
  top_left_cell.font  = Font(b=True, color="000000")
  top_left_cell.alignment = Alignment(horizontal="center", vertical="center")
  seller_cell = ws['A3']
  seller_cell.font  = Font(b=True, color="000000")
  seller_cell.alignment = Alignment(horizontal="center", vertical="center")
  # Put the borders
  thin = Side(border_style="thin", color="000000")
  for col in range(1, column_count+1):
    for row in range(1, row_count):
      cell = ws.cell(row=row, column=col)
      cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
  # Saving document
  wb.save(name)

