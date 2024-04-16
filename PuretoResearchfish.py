import pandas as pd
import os
from win32com import client as wc
import numpy as np

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

    # Save the workbook
    wb.Save()

    # Close the workbook and quit Excel
    wb.Close()
    excel.Quit()

    # Read the Excel file into a DataFrame
    df = pd.read_excel(temp_file)

    # Define a function to handle 'N/A' or 'n/a' as null values
    def na_handler(x):
        if isinstance(x, str) and x.lower() in ['n/a', 'na']:
            return np.nan
        return x

    # Apply the 'na_handler' function to the DataFrame
    df = df.apply(lambda x: x.map(na_handler) if x.name == "Funder Project Reference" else x)

    # Find the column with header "Funder Project Reference"
    funder_proj_ref_header = "Funder Project Reference"
    if funder_proj_ref_header not in df.columns:
        print(f"Error: Column '{funder_proj_ref_header}' not found.")
        return

    # Remove rows with blank "Funder Project Reference" (including 'N/A' or 'n/a')
    df = df.dropna(subset=[funder_proj_ref_header], how='any')

    # Save the DataFrame to a new Excel file
    df.to_excel(output_file, index=False)

    # Delete the temporary file
    os.remove(temp_file)

    print("Duplicate rows and rows with blank 'Funder Project Reference' (including 'N/A' or 'n/a') removed. Output saved to:", output_file)

if __name__ == "__main__":
    # Prompt user for input file path
    input_file = input("Enter the path to the input Excel file (.xlsx): ")

    # Prompt user for output file path
    output_file = input("Enter the path for the output Excel file (.xlsx): ")

    # Call the function to remove duplicates using Excel and handle "Funder Project Reference" with pandas
    remove_duplicates_and_empty_rows(input_file, output_file)
