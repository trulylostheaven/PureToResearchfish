import pandas as pd

def remove_duplicate_rows(input_file, output_file):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(input_file)
    
    # Remove rows where all columns are the same
    df = df.drop_duplicates(subset=None, keep="first")
    
    # Write the modified DataFrame back to a new Excel file
    df.to_excel(output_file, index=False)

if __name__ == "__main__":
    # Prompt user for input file path
    input_file = input("Enter the path to the input Excel file (.xlsx): ")
    
    # Prompt user for output file path
    output_file = input("Enter the path for the output Excel file (.xlsx): ")
    
    # Call the function to remove duplicate rows
    remove_duplicate_rows(input_file, output_file)
    print("Duplicate rows removed. Output saved to:", output_file)
