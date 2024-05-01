# PureToResearchfish

Please note that this code is currently under development. It is prone to errors and the code is still revising. This may not work for your institution so please review what the code does before using it.
PureToResearchfish is a Python program designed to streamline the manipulation of data extracted from Elsevier Pure, making it easier to upload to Researchfish.

## Requirements
- Python
- Excel
- pandas
- numpy
- fuzzywuzzy

## Features
**PuretoResearchfish.py**
1. **Delete Duplicates**: Removes duplicate rows from the dataset.
2. **Filter by "Funder Project Reference"**:
   - Removes rows with blank "Funder Project Reference" fields.
   - Splits "specific" rows with more than one "Funder Project Reference" id.
4. **Filter by DOIs and Additional Source IDs**:
   - Keeps rows with a DOI OR where Additional Source IDs start with "PubMed:".
   - Removes the "PubMed:" prefix from Additional Source IDs.
5. **Clear Additional Source IDs if DOIs are Present**: Clears the "Additional Source IDs" field if a DOI is present.
6. **Filter by "Funder Project Reference"**: Removes rows with just dates or via institution in "Funder Project Reference"

**PureAddFunder**
1. **Match Name to Funder Organisation**: Adds Funder Organisation while matching to Name.
2. **Add Funder ID**: Adds the Funder ID matching to Funder Organisation.

## Usage
1. Clone the repository:
   ```bash
   git clone https://github.com/your-username/PureToResearchfish.git
   cd PureToResearchfish
2. Install the required Python packages:
   ```bash
   pip install pandas numpy
   pip install fuzzywuzzy
3. Run the program:
   ```bash
   python run_program.py
4. Run program PuretoResearchfish.py (this will run both programs in order)
5. Select input_file.xlsx, select comparison_xlsx, and select comparison_sheet

## How it Works
1. **Delete Duplicates**
   The program utilizes Excel to remove duplicate rows from the dataset. This is run twice, once in the beginning and at the end.
2. **Filter by "Funder Project Reference"**
   Rows with blank "Funder Project Reference" fields are removed using pandas.
3. **Filter by DOIs and Additional Source IDs**
   The program filters rows based on the presence of DOIs (Digital Object Identifiers) or "PubMed:" in Additional Source IDs.
4. **Clear Additional Source IDs if DOIs are Present**
   If a DOI is found in a row, the program clears the corresponding "Additional Source IDs".
