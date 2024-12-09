import os
import pandas as pd
from fuzzywuzzy import process
import tkinter as tk
from tkinter import Tk, filedialog, ttk
import logging
from utils import get_unique_output_file

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

def select_files():
    root = Tk()
    root.withdraw()
    input_file = filedialog.askopenfilename(title="Select Input File", filetypes=[("Excel files", "*.xlsx")])
    comparison_file = filedialog.askopenfilename(title="Select Comparison File", filetypes=[("Excel files", "*.xlsx")])
    if not input_file or not comparison_file:
        logging.error("File selection canceled. Exiting...")
        raise SystemExit("File selection was not completed.")
    return input_file, comparison_file

def select_sheet(comparison_file):
    with pd.ExcelFile(comparison_file) as xls:
        sheet_names = xls.sheet_names

    root = tk.Tk()
    root.title("Select Sheet")

    selected_sheet = tk.StringVar()
    
    def on_select():
        selected = sheet_listbox.curselection()
        if selected:
            selected_sheet.set(sheet_names[selected[0]])
            root.quit()
            root.destroy()

    ttk.Label(root, text="Available Sheets:").pack()
    sheet_listbox = tk.Listbox(root, selectmode=tk.SINGLE)
    for name in sheet_names:
        sheet_listbox.insert(tk.END, name)
    sheet_listbox.pack()

    ttk.Button(root, text="Select", command=on_select).pack()
    root.mainloop()

    if not selected_sheet.get():
        raise ValueError("No sheet selected. Please try again.")
    return selected_sheet.get()


def fuzzy_match_name(name, comparison_names, threshold=90):
    result = process.extractOne(name, comparison_names)
    if result is None:
        return None
    matched_name, score = result
    return matched_name if score >= threshold else None

def process_files(input_file, comparison_file):
    # Select the sheet from the comparison file
    logging.info("Prompting user to select a sheet from the comparison file...")
    selected_sheet = select_sheet(comparison_file)

    # Load input and selected comparison sheet
    logging.info(f"Loading data from selected sheet: {selected_sheet}")
    input_df = pd.read_excel(input_file)
    with pd.ExcelFile(comparison_file) as xls:
        comparison_df = pd.read_excel(xls, sheet_name=selected_sheet)

    # Clean and validate comparison DataFrame
    comparison_df.columns = comparison_df.columns.str.strip()
    if "Name" not in comparison_df.columns or "RF Funder ID" not in comparison_df.columns:
        raise ValueError("The comparison file must contain 'Name' and 'RF Funder ID' columns.")

    comparison_names = comparison_df["Name"].dropna().astype(str).unique().tolist()
    name_to_funder_id = dict(zip(comparison_df["Name"], comparison_df["RF Funder ID"]))

    # Add and fill new columns
    logging.info("Performing fuzzy matching and filling data...")
    input_df["Matched Name"] = input_df["Funding organisation(s)"].fillna("").astype(str).map(
        lambda x: fuzzy_match_name(x, comparison_names, threshold=90)
    )
    input_df["RF Funder ID"] = input_df["Matched Name"].map(name_to_funder_id)

    # Drop rows with unmatched names
    input_df_cleaned = input_df.dropna(subset=["Matched Name"])

    # Rename "Matched Name" to "Name"
    input_df_cleaned = input_df_cleaned.rename(columns={"Matched Name": "Name"})

    # Reorder columns to place "Name" as the last column
    columns = [col for col in input_df_cleaned.columns if col != "Name"] + ["Name"]
    input_df_cleaned = input_df_cleaned[columns]

    return input_df_cleaned

def save_output(df, input_file):
    output_file = get_unique_output_file(input_file, suffix="_processed")
    logging.info(f"Saving processed data to {output_file}...")
    df.to_excel(output_file, index=False)
    return output_file

def main():
    try:
        # File selection
        input_file, comparison_file = select_files()

        # Process files with sheet selection and data matching
        processed_df = process_files(input_file, comparison_file)

        # Save the output with a unique file name
        output_file = save_output(processed_df, input_file)
        logging.info(f"Processing complete. File saved at: {output_file}")

    except Exception as e:
        logging.error(f"An error occurred: {e}")
        raise
    
if __name__ == "__main__":
    main()
