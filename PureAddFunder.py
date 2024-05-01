import os
import pandas as pd
import tkinter as tk
from tkinter import Tk, filedialog, ttk
from fuzzywuzzy import process
import win32com.client as wc

class ExcelFileHandler:
    def __init__(self, file_path):
        self.file_path = file_path
        self.excel = wc.Dispatch("Excel.Application")
        self.excel.Visible = False
        self.wb = None

    def __enter__(self):
        self.wb = self.excel.Workbooks.Open(os.path.abspath(self.file_path))
        return self.wb.ActiveSheet

    def __exit__(self, exc_type, exc_value, traceback):
        if self.wb:
            self.wb.Save()
            self.wb.Close()
        self.excel.Quit()

def add_columns(input_file, output_file):    
    # Read the Excel file into a DataFrame
    with pd.ExcelFile(input_file) as xls:
        df = pd.read_excel(xls)

    # Add new columns
    df['RF Funder ID'] = ''

    # Save the DataFrame to Excel
    df.to_excel(output_file, index=False)

    return df

def fuzzy_match_name(name, comparison_names, threshold=90):
    # Perform fuzzy matching to find the closest match
    matched_name, score = process.extractOne(name, comparison_names)
    if score >= threshold:
        return matched_name
    else:
        return None

def select_comparison_sheet(comparison_file):
    root = tk.Tk()
    root.title("Select Sheet")

    sheet_label = ttk.Label(root, text="Available sheet names in the comparison file:")
    sheet_label.pack()

    sheet_listbox = tk.Listbox(root, selectmode=tk.SINGLE)
    sheet_listbox.pack()

    def on_select():
        selected_index = sheet_listbox.curselection()
        if selected_index:
            selected_sheet.set(sheet_listbox.get(selected_index))
            root.quit()
            root.destroy()

    with pd.ExcelFile(comparison_file) as xls:
        sheet_names = xls.sheet_names

    for name in sheet_names:
        sheet_listbox.insert(tk.END, name)

    select_button = ttk.Button(root, text="Select", command=on_select)
    select_button.pack()

    selected_sheet = tk.StringVar()
    root.mainloop()

    return selected_sheet.get()

def compare_and_fill(df, comparison_file, output_file):
    selected_sheet = select_comparison_sheet(comparison_file)
    with pd.ExcelFile(comparison_file) as xls:
        comparison_df = pd.read_excel(xls, sheet_name=selected_sheet)

    # Remove leading and trailing whitespace from column names
    comparison_df.columns = comparison_df.columns.str.strip()

    print("Columns in the comparison DataFrame after stripping whitespace:")
    print(comparison_df.columns)

    # Extract unique names from the comparison DataFrame
    if "Name" not in comparison_df.columns:
        raise ValueError("The 'Name' column does not exist in the comparison DataFrame.")

    comparison_names = comparison_df["Name"].unique()

    # Apply fuzzy matching to each row in the input DataFrame
    df["Name"] = df["Funding organisation(s)"].apply(lambda x: fuzzy_match_name(x, comparison_names, threshold=90))

    # Save the DataFrame to Excel
    df.to_excel(output_file, index=False)

    return df, selected_sheet

def fill_rf_funder_id(df, comparison_df, selected_sheet, output_file):
    # Rename the column in the comparison DataFrame to match the input DataFrame
    comparison_df = comparison_df.rename(columns={" Name": "Name"})
    
    # Create a dictionary mapping names to RF Funder IDs from the comparison DataFrame
    name_to_funder_id = dict(zip(comparison_df["Name"], comparison_df["RF Funder ID"]))
    
    # Fill in RF Funder ID based on the matched names
    df["RF Funder ID"] = df["Name"].map(name_to_funder_id)
    
    # Save the DataFrame to Excel
    df.to_excel(output_file, index=False)

    return df

def clean_up_data(df, output_file):
    # Drop rows with missing values in the "Name" column
    df_cleaned = df.dropna(subset=["Name"])
    
    # Save the cleaned DataFrame to Excel
    df_cleaned.to_excel(output_file, index=False)

    return df_cleaned

def main():
    # Use tkinter to prompt user to choose input and comparison files
    root = Tk()
    root.withdraw()

    input_file = filedialog.askopenfilename(initialdir=os.getcwd(), title="Select Input File", filetypes=[("Excel files", "*.xlsx")])
    comparison_file = filedialog.askopenfilename(initialdir=os.getcwd(), title="Select Comparison File", filetypes=[("Excel files", "*.xlsx")])

    if not input_file or not comparison_file:
        print("One or both files not selected. Exiting...")
        return

    # Determine the output file name
    base_output_file = os.path.splitext(input_file)[0] + "-comparison"
    output_file = f"{base_output_file}.xlsx"
    output_counter = 1

    # Check if the output file already exists, if yes, iterate the counter
    while os.path.exists(output_file):
        output_counter += 1
        output_file = f"{base_output_file}{output_counter}.xlsx"

    print("Input file:", input_file)
    print("Comparison file:", comparison_file)
    print("Output file:", output_file)

    # Call the function to add columns
    df = add_columns(input_file, output_file)

    # Call the function to compare and fill data from the comparison file
    result_df, selected_sheet = compare_and_fill(df, comparison_file, output_file)

    # Read the comparison file into a DataFrame
    with pd.ExcelFile(comparison_file) as xls:
        comparison_df = pd.read_excel(xls, sheet_name=selected_sheet)

    # Call the function to fill RF Funder IDs based on the comparison data
    filled_df = fill_rf_funder_id(result_df, comparison_df, selected_sheet, output_file)

    # Call the function to clean up the data by removing rows with missing "Name" values
    cleaned_df = clean_up_data(filled_df, output_file)

if __name__ == "__main__":
    main()
