# PureToResearchfish
A simple python program that manipulates the data pulled from Elsevier Pure to make it easier to upload to Researchfish.

# Progress
1. Deletes Duplicates
2. Removes all rows with blank Funder Project Reference
3. Keeps only rows where there is a DOI OR Additional Source IDs starts with PubMed:
4. Removes "PubMed:" string from cells in Additional Source IDs leaving only ID
5. Keeps only Additional Source IDs when there is no DOI
