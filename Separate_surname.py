import pandas as pd
import os
from dotenv import load_dotenv

# Load environment variables from the custom env file
load_dotenv(dotenv_path='separate_surname.env')

# Get values from the env file
file_path = os.getenv('INPUT_FILE_PATH')
output_file = os.getenv('OUTPUT_FILE_PATH')
sheet_name = os.getenv('SHEET_NAME')
column_name = os.getenv('COLUMN_NAME')

# Read the Excel file
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Extract the last word from the full name as the last name
df['Last Name'] = df[column_name].apply(lambda x: str(x).split()[-1] if pd.notnull(x) else '')

# Save the modified DataFrame to a new Excel file
df.to_excel(output_file, index=False)

print(f"Last names extracted and saved to {output_file}")