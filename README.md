# Excel Contacts Flattener

This script transforms an Excel file containing multiple contact rows per company into a CSV file with a single, "flattened" row per company.

The company information is preserved once, while contact details are transposed into repeated, suffixed columns (e.g., `Nom_1`, `Email pro_1`, `Nom_2`, `Email pro_2`, etc.) on the same row. The number of contact blocks is determined by the company with the most associated contacts in the input file.

## How It Works

1.  Reads the first sheet of the input Excel file.
2.  Groups all rows by the company name (`Raison Sociale`).
3.  Determines the maximum number of contacts associated with a single company.
4.  Dynamically generates new column headers for the flattened contact information.
5.  Constructs a new data structure with one row per company.
6.  Writes the result to a CSV file, encoded in UTF-8 with BOM and using a semicolon (`;`) as the separator.

## Input File Structure

The script expects an Excel file (`.xlsx`) with the following columns in the first sheet:

*   `Raison Sociale`
*   `Description d'activité`
*   `Population`
*   `Localisation (departement)`
*   `Civilité`
*   `Nom`
*   `Prénom`
*   `Fonction`
*   `Email pro`
*   `Localisation`
*   `Date de Création`
*   `CEO`
*   `SIREN`
*   `SIRET`
*   `Descriptif NAF`

## Prerequisites

This script requires Python and the libraries listed in `requirements.txt`.

Install the dependencies using pip:
```bash
pip install -r requirements.txt
```

## Usage

Run the script from your terminal, providing the input Excel file path and the desired output CSV file path as arguments.

**Command:**
```bash
python main.py <input_file.xlsx> <output_file.csv>
```

**Example:**
```bash
python main.py contacts_source.xlsx contacts_flattened.csv
```

This will process `contacts_source.xlsx` and generate a new file named `contacts_flattened.csv` in the same directory.
