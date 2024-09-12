
Here's your script with the new comment section at the top, including the version, last updated date, and a brief description:

python
Skopiuj kod
# ------------------------------------------------------------------------------
# Version: 1.6
# Last updated: 2024-09-12
# Description: 
# This script generates an Excel workbook containing Data Warehouse and 
# Data Mart objects. It applies formatting such as text alignment, column 
# width adjustments, and adds borders around cells. It also auto-adjusts 
# the row heights for descriptions and includes hyperlinks for certain 
# columns. The script reads data from two dataframes representing updated 
# tables for Data Warehouse and Data Mart objects and exports the results 
# to a new Excel file.
# ------------------------------------------------------------------------------

import pandas as pd
import openpyxl
from openpyxl.styles import Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from datetime import datetime

def adjust_column_widths(ws):
    """Auto-adjust column widths for columns B and C (C is 8x B)."""
    column_b_width = 10  # Adjust width of column B
    column_c_width = column_b_width * 8  # Column C should be 8x width of column B

    ws.column_dimensions['B'].width = column_b_width
    ws.column_dimensions['C'].width = column_c_width

def apply_full_borders(ws):
    """Apply thin borders around all cells in the sheet."""
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    for row in ws.iter_rows():
        for cell in row:
            cell.border = thin_border

def apply_formatting(ws):
    """Apply specific formatting for each column."""
    for row in ws.iter_rows():
        for cell in row:
            col_letter = get_column_letter(cell.column)

            if col_letter == 'A':  # Column A (Object number) should be bottom-aligned
                cell.alignment = Alignment(horizontal='left', vertical='bottom')

            elif col_letter == 'B':  # Column B (Object name) should be hyperlink format
                cell.alignment = Alignment(horizontal='left', vertical='bottom')
                # Apply hyperlink style
                cell.font = Font(color="0000FF", underline="single")

            elif col_letter == 'C':  # Column C (Description) should be center-aligned both horizontally and vertically
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

            else:  # Other columns should be bottom-aligned
                cell.alignment = Alignment(horizontal='left', vertical='bottom')

def create_and_format_workbook(output_file, updated_table1_df, updated_table2_df):
    """Create new workbook and write data with formatting."""
    wb_new = openpyxl.Workbook()
    ws = wb_new.active
    ws.title = 'object list'

    # Define bold font for headers with increased size
    header_font = Font(bold=True, size=12)

    # Write Data Warehouse headers and format
    dwh_headers = ['Dwh object number', 'Dwh object name', 'Dwh object description', 'Reports Tag', 'checked']
    for col_idx, header in enumerate(dwh_headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font

    # Write Data Warehouse objects
    for r_idx, row in enumerate(updated_table1_df.itertuples(index=False), 2):
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            if c_idx == 2:  # Column B: Hyperlink format
                ws.cell(row=r_idx, column=2).hyperlink = f"#{row[1]}!A1"  # Add hyperlink
                ws.cell(row=r_idx, column=2).style = "Hyperlink"  # Default hyperlink formatting
            if c_idx == 3:  # Adjust row height for descriptions in column C
                ws.row_dimensions[r_idx].height = 45  # Adjust height

    # Write Data Mart headers and objects
    dm_start_row = len(updated_table1_df) + 3
    dm_headers = ['Data mart object number', 'Data mart object name', 'Data mart object description', 'Reports Tag']
    for col_idx, header in enumerate(dm_headers, 1):
        cell = ws.cell(row=dm_start_row - 1, column=col_idx, value=header)
        cell.font = header_font

    for r_idx, row in enumerate(updated_table2_df.itertuples(index=False), dm_start_row):
        for c_idx, value in enumerate(row, 1):
            ws.cell(row=r_idx, column=c_idx, value=value)

    # Apply column width adjustments
    adjust_column_widths(ws)

    # Apply full borders around all cells
    apply_full_borders(ws)

    # Apply text alignment and formatting
    apply_formatting(ws)

    # Save the new workbook
    wb_new.save(output_file)
    print(f"Workbook created and saved at {output_file}")

def main():
    input_file = r"C:\Users\pttom\OneDrive\Pulpit\DWH_Source_Matrix (2).xlsx"
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_file = rf"C:\Users\pttom\OneDrive\Pulpit\Object_List_Only_{timestamp}.xlsx"

    # Example data
    updated_table1_df = pd.DataFrame({
        'Dwh object number': [1, 2, 3],
        'Dwh object name': ["dwh.object1", "dwh.object2", "dwh.object3"],
        'Dwh object description': ["Short description", "Long description " * 5, "Another description"],
        'Reports Tag': ["Tag1", "Tag2", "Tag3"],
        'checked': [True, False, True]
    })

    updated_table2_df = pd.DataFrame({
        'Data mart object number': [1, 2, 3],
        'Data mart object name': ["dm.object1", "dm.object2", "dm.object3"],
        'Data mart object description': ["Short description for Data Mart", "Long description " * 3, "Another Data Mart description"],
        'Reports Tag': ["Tag1", "Tag2", "Tag3"],
        'checked': [True, False, True]
    })

    # Create and save the formatted workbook
    create_and_format_workbook(output_file, updated_table1_df, updated_table2_df)

if __name__ == "__main__":
    main()


