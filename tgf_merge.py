import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Border, Side

# File path
file_path = '/Users/marvin/acn/sw-web-automation/TGF/Test/TGF_01_0014_before_SAP.xlsx'

# Read the Excel file and the two sheets
df1 = pd.read_excel(file_path, sheet_name='1')
df2 = pd.read_excel(file_path, sheet_name='2')

# Identify common and unique columns
common_columns = df1.columns.intersection(df2.columns)
unique_columns_df2 = df2.columns.difference(df1.columns)

# Merge the DataFrames
merged_df = pd.concat([df1, df2[common_columns]], ignore_index=True)

# Add unique columns from df2 to merged_df
for col in unique_columns_df2:
    merged_df[col] = df2[col]

# Save the merged DataFrame to a new Excel file
output_path = '/Users/marvin/acn/sw-web-automation/TGF/Test/merged_output.xlsx'
merged_df.to_excel(output_path, index=False)

# Load the original workbook and the new workbook
original_wb = load_workbook(file_path)
original_ws = original_wb['1']
new_wb = load_workbook(output_path)
new_ws = new_wb.active

# Define a border style
thin_border = Border(left=Side(style='thin'), 
                     right=Side(style='thin'), 
                     top=Side(style='thin'), 
                     bottom=Side(style='thin'))

# Copy the font and style from the original first sheet to the new sheet
for row in new_ws.iter_rows(min_row=1, max_row=new_ws.max_row, min_col=1, max_col=new_ws.max_column):
    for cell in row:
        original_cell = original_ws[cell.coordinate] if cell.coordinate in original_ws else None
        if original_cell:
            cell.font = original_cell.font.copy()
            cell.alignment = original_cell.alignment.copy()
            cell.border = original_cell.border.copy()
            cell.fill = original_cell.fill.copy()
        cell.border = thin_border  # Apply the border to all cells

# Adjust column widths
for col in new_ws.columns:
    max_length = 0
    column = col[0].column_letter  # Get the column name
    for cell in col:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    new_ws.column_dimensions[column].width = adjusted_width

# Save the adjusted workbook
new_wb.save(output_path)