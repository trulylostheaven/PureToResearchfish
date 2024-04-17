import os
import pandas as pd
import numpy as np
import win32com.client as wc
import re

def remove_duplicates_from_excel(input_file, temp_file):
    # Check if temp_file already exists
    if not os.path.exists(temp_file):
        # Copy the input file to a temporary file
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

def handle_funder_project_reference(input_file, output_file):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(input_file)

    # Define a function to handle 'N/A' as null values
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

    print("Duplicate rows and rows with blank 'Funder Project Reference' (including 'N/A' or 'n/a') removed. Output saved to:", output_file)

def filter_by_dois_and_additional_ids(input_file, output_file):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(input_file)

    # Fill NaN values in "Additional source IDs" column with an empty string
    df["Additional source IDs"] = df["Additional source IDs"].fillna("")

    # Create a mask for rows where "DOIs (Digital Object Identifiers)" is not NaN
    dois_not_null_mask = df["DOIs (Digital Object Identifiers)"].notnull()

    # Create a mask for rows where "Additional source IDs" contains "PubMed:"
    pubmed_mask = df["Additional source IDs"].str.contains("PubMed:", na=False)

    # Combine both conditions to keep rows where DOIs are not null or where Additional source IDs contain "PubMed:"
    df_filtered = df[dois_not_null_mask | pubmed_mask].copy()  # Make a copy to avoid the SettingWithCopyWarning

    # Remove "PubMed:" prefix from "Additional source IDs" column
    df_filtered.loc[pubmed_mask, "Additional source IDs"] = df_filtered.loc[pubmed_mask, "Additional source IDs"].str.replace(r'^\s*PubMed:\s*', '', regex=True)

    # Save the DataFrame to a new Excel file
    df_filtered.to_excel(output_file, index=False)

    print("Filtered rows based on 'DOIs (Digital Object Identifiers)' and 'Additional source IDs'. Output saved to:", output_file)

def clear_additional_ids_if_doi_present(input_file, output_file):
    # Read the Excel file
    df = pd.read_excel(input_file)

    # Check if "DOIs (Digital Object Identifiers)" column has data
    dois_column = "DOIs (Digital Object Identifiers)"
    additional_ids_column = "Additional source IDs"

    # Mask for rows where DOIs are present
    dois_present_mask = df[dois_column].notnull()

    # Clear "Additional source IDs" where DOIs are present
    df.loc[dois_present_mask, additional_ids_column] = np.nan

    # Save the modified DataFrame to a new Excel file
    df.to_excel(output_file, index=False)
    print("Cleared 'Additional source IDs' where 'DOIs (Digital Object Identifiers)' are present. Output saved to:", output_file)

def remove_rows_with_dates_or_via(input_file, output_file):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(input_file)

    # Function to check if a value is in the format "##/##/##"
    def is_date_format(value):
        if not isinstance(value, str):
            return False
        parts = value.split('/')
        if len(parts) != 3:
            return False
        try:
            day, month, year = map(int, parts)
            if 1 <= day <= 31 and 1 <= month <= 12:
                return True
        except ValueError:
            pass
        return False

    # Function to check if a value contains "Via [Institution Name]" without numbers
    def contains_via_institution(value):
        if not isinstance(value, str):
            return False
        pattern = r"via [a-zA-Z\s]+"
        return bool(re.match(pattern, value, re.IGNORECASE))

    # Function to check if a value contains a removal keyword
    def contains_keyword(value):
        if not isinstance(value, str):
            return False
        keywords = [
            "COVID 19 Supplement",
            "Researcher let 1",
            "Diana Tay",
            "Amendment #5",
            "Amendment #6"
        ]
        return any(keyword in value for keyword in keywords)

    # Create boolean masks for dates, "Via [Institution Name]", and keywords
    date_mask = df["Funder Project Reference"].astype(str).apply(is_date_format)
    via_mask = df["Funder Project Reference"].astype(str).apply(contains_via_institution)
    keyword_mask = df["Funder Project Reference"].astype(str).apply(contains_keyword)

    # Combine the masks using logical OR to identify rows to remove
    rows_to_remove = date_mask | via_mask | keyword_mask

    # Keep rows where the value is not a date, does not contain "Via [Institution Name]", 
    # or a hard-coded removal keyword
    df = df[~rows_to_remove]

    # Write the updated DataFrame to a new Excel file
    df.to_excel(output_file, index=False)

    print("Rows with dates, 'Via [Institution Name]', 'COVID 19 Supplement', 'Amendment #5', or 'Amendment #6' in 'Funder Project Reference' column removed. Output saved to:", output_file)

def remove_via_notes_from_funder_reference(input_file, output_file):
    # Read the Excel file
    df = pd.read_excel(input_file)

    # Define a function to remove via notes
    def remove_via_notes(cell_value):
        if isinstance(cell_value, str):
            # Updated regular expression to remove "(via [institution name])" or "via [institution name]"
            cell_value = re.sub(r'\s*\([^()]*via\s+[^\s()]+\s*[^()]*\)|\s*via\s+[^\s()]+\s*', '', cell_value, flags=re.IGNORECASE)
            # Remove any remaining parentheses
            cell_value = re.sub(r'\([^()]*\)', '', cell_value)
        return cell_value

    # Apply the function to the "Funder Project Reference" column
    df['Funder Project Reference'] = df['Funder Project Reference'].apply(remove_via_notes)

    # Save the modified DataFrame to a new Excel file
    df.to_excel(output_file, index=False)

    print("Via notes removed from 'Funder Project Reference' column. Output saved to:", output_file)
    
def main():
    # Prompt user for input file path
    input_file = input("Enter the path to the input Excel file (.xlsx): ")

    # Prompt user for output file path
    output_file = input("Enter the path for the output Excel file (.xlsx): ")

    # Create a temporary file path
    temp_file = "temp_file.xlsx"

    # Call the function to remove duplicates using Excel
    remove_duplicates_from_excel(input_file, temp_file)

    # Call the function to handle "Funder Project Reference" header using pandas
    handle_funder_project_reference(temp_file, temp_file)

    # Call the function to filter rows based on "DOIs (Digital Object Identifiers)" and "Additional source IDs"
    filter_by_dois_and_additional_ids(temp_file, temp_file)
    
    # Call the new function to clear "Additional source IDs" where "DOIs" are present
    clear_additional_ids_if_doi_present(temp_file, temp_file)
    
    # Call the function to remove rows with dates in "Funder Project Reference" column
    remove_rows_with_dates_or_via(temp_file, temp_file)

    # Call the new function to remove via notes from "Funder Project Reference" column
    remove_via_notes_from_funder_reference(temp_file, temp_file)

    # Remove duplicates from the final output file
    remove_duplicates_from_excel(temp_file, output_file)

    # Delete the temporary file
    os.remove(temp_file)

if __name__ == "__main__":
    main()
