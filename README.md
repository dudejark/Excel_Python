# 🧾 Extract Last Names from Excel

This Python script helps you extract last names (surnames) from a full name column in an Excel file. It's especially useful for data cleaning or splitting names into structured formats.

## 📄 What It Does

Given an Excel file with a column of full names (e.g., "John Doe", "Amitabh Bachchan"), the script:
- Reads the file using `pandas`
- Splits each full name
- Extracts the last word as the **surname**
- Writes the result to a new column called **`Last Name`**
- Saves the updated Excel file

## 🧰 Requirements

- Python 3.x
- pandas
- openpyxl

Install the requirements using:

```bash
pip install pandas openpyxl
