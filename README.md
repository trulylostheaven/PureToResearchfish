# PureToResearchfish

PureToResearchfish is a Python program designed to streamline the manipulation of data extracted from Elsevier Pure, making it easier to upload to Researchfish.

## Requirements
- Python
- Excel
- pandas
- numpy

## Features
1. **Delete Duplicates**: Removes duplicate rows from the dataset.
2. **Filter by "Funder Project Reference"**: Removes rows with blank "Funder Project Reference" fields.
3. **Filter by DOIs and Additional Source IDs**:
   - Keeps rows with a DOI OR where Additional Source IDs start with "PubMed:".
   - Removes the "PubMed:" prefix from Additional Source IDs.
4. **Clear Additional Source IDs if DOIs are Present**: Clears the "Additional Source IDs" field if a DOI is present.

## Usage
1. Clone the repository:
   ```bash
   git clone https://github.com/your-username/PureToResearchfish.git
   cd PureToResearchfish
2. Install the required Python packages:
   ```bash
   pip install pandas numpy
3. Run the program:
   ```bash
   python PureToResearchfish.py
4. You will be prompted to input the paths to the input and output Excel files:
   ```bash
   Enter the path to the input Excel file (.xlsx): path/to/input/file.xlsx
   Enter the path for the output Excel file (.xlsx): path/to/output/file.xlsx
