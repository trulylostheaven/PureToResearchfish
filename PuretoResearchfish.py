import pandas as pd
import os
from win32com import client as wc

def remove_duplicates_with_excel(input_file, output_file):
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
    
    # Call the function to remove duplicates using Excel
    remove_duplicates_with_excel(input_file, output_file)
    print("Duplicate rows removed using Excel. Output saved to:", output_file)
