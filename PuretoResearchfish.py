import pandas as pd
import os
from win32com import client as wc

def remove_duplicates_and_empty_rows(input_file, output_file):
    # Copy the input file to a temporary file
    temp_file = "temp_file.xlsx"
    os.system(f"copy \"{input_file}\" \"{temp_file}\"")

    # Create an Excel Application
    excel = wc.Dispatch("Excel.Application")
    excel.Visible = False

    # Open the temporary file
    wb = excel.Workbooks.Open(os.path.abspath(temp_file))

    # Select the active sheet
    ws = wb.ActiveSheet

    # Remove duplicates from the entire used range
    ws.UsedRange.RemoveDuplicates(Columns=range(1, ws.UsedRange.Columns.Count + 1), Header=1)

    # Find the index of the column with header "Funder Project Reference"
    funder_proj_ref_header = "Funder Project Reference"
    funder_proj_ref_col_index = None
    for col in range(1, ws.UsedRange.Columns.Count + 1):
        header_value = ws.Cells(1, col).Value
        if header_value == funder_proj_ref_header:
            funder_proj_ref_col_index = col
            break

    if funder_proj_ref_col_index is None:
        print(f"Error: Column '{funder_proj_ref_header}' not found.")
        wb.Close()
        excel.Quit()
        os.remove(temp_file)
        return

    # Loop through rows to check and delete rows with blank "Funder Project Reference"
    rows_to_delete = []
    for row in range(2, ws.UsedRange.Rows.Count + 1):  # Start from 2nd row assuming there's a header
        cell_value = ws.Cells(row, funder_proj_ref_col_index).Value
        if cell_value is None:
            rows_to_delete.append(row)

    # Delete the rows with blank "Funder Project Reference"
    for row in reversed(rows_to_delete):  # Deleting in reverse order to avoid shifting indexes
        ws.Rows(row).Delete()

    # Save the workbook
    wb.SaveAs(os.path.abspath(output_file), FileFormat=51)

    # Close the workbook and quit Excel
    wb.Close()
    excel.Quit()

    # Delete the temporary file
    os.remove(temp_file)

if __name__ == "__main__":
    # Prompt user for input file path
    input_file = input("Enter the path to the input Excel file (.xlsx): ")

    # Prompt user for output file path
    output_file = input("Enter the path for the output Excel file (.xlsx): ")

    # Call the function to remove duplicates and empty rows using Excel
    remove_duplicates_and_empty_rows(input_file, output_file)
    print("Duplicate rows and rows with blank 'Funder Project Reference' removed using Excel. Output saved to:", output_file)
