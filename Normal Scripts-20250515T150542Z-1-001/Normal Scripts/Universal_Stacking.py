"""
This script is used to stack the latest_website_url
"""

import pandas as pd

# Define the delimiters
delimiters = [","]

# Load the input Excel file
input_file = r"C:\Users\Kompass\Desktop\Test Files\Phones_1.3.2025.xlsx"  # Replace with your input file path
df = pd.read_excel(input_file)

# Function to split and stack rows
def split_and_stack(df, column, delimiter):
    # Split the column by the delimiter and expand into multiple rows
    df = df.assign(**{column: df[column].str.split(delimiter)}).explode(column)
    # Strip any leading/trailing whitespace from the split values
    df[column] = df[column].str.strip()
    return df

# Process each column that may contain delimited values
stacked_df = df.copy()
for column in ["Phones"]:  # Add more columns if needed
    for delimiter in delimiters:
        # Check if the delimiter exists in the column
        if stacked_df[column].astype(str).str.contains(delimiter).any():
            # Apply the splitting and stacking function
            stacked_df = split_and_stack(stacked_df, column, delimiter)
            break  # Stop after the first matching delimiter

# Save the output to a new Excel file
output_file = r"C:\Users\Kompass\Desktop\Test Files\Stacking_Phones_1.3.2025.xlsx"  # Replace with your desired output file path
stacked_df.to_excel(output_file, index=False)

print(f"Output saved to {output_file}")


"""
This script is used to stack the latest_website_url and remove ?trk..

"""








import pandas as pd

# Define the delimiters
delimiters = [",", "*", ";", "/", " "]

# Load the input Excel file
input_file = r"C:\Users\Kompass\Downloads\Import Export Data_4April025.xlsx" # Replace with your input file path
df = pd.read_excel(input_file)

# Function to split and stack rows
def split_and_stack(df, column, delimiter):
    # Convert column to string type first
    df[column] = df[column].astype(str)
    # Split the column by the delimiter and expand into multiple rows
    df = df.assign(**{column: df[column].str.split(delimiter)}).explode(column)
    # Strip any leading/trailing whitespace from the split values
    df[column] = df[column].str.strip()
    return df

# Function to clean URLs by removing query parameters
def clean_url(url):
    if pd.isna(url) or url == 'nan':
        return url
    # Split the URL at "?" and keep only the part before it
    return str(url).split("?")[0]

# Process each column that may contain delimited values
stacked_df = df.copy()
for column in ["EMAIL"]:  # Add more columns if needed
    # Convert to string type first
    stacked_df[column] = stacked_df[column].astype(str)
    for delimiter in delimiters:
        # Check if the delimiter exists in the column
        if stacked_df[column].str.contains(delimiter).any():
            # Apply the splitting and stacking function
            stacked_df = split_and_stack(stacked_df, column, delimiter)
            break  # Stop after the first matching delimiter

# Clean the URLs in the Email column
stacked_df["EMAIL"] = stacked_df["EMAIL"].apply(clean_url)

# Remove rows with 'nan' (converted from NaN) if needed
stacked_df = stacked_df[stacked_df["EMAIL"] != 'nan']

# Save the output to a new Excel file
output_file = r"C:\Users\Kompass\Downloads\Stack_Email_Import Export Data_4April025.xlsx"  # Replace with your desired output file path
stacked_df.to_excel(output_file, index=False)

print(f"Output saved to {output_file}")










import pandas as pd

# Read data from the sheet named 'Sheet2'
FILEPATH = r"C:\Users\Kompass\Downloads\Email Address_Indian Exporters & Importers.xlsx"
df = pd.read_excel(FILEPATH, sheet_name='Sheet1', engine='calamine')

# Debugging: Print available columns
print("Available columns:", df.columns)

# Standardizing column names
df.columns = df.columns.str.strip()  # Remove extra spaces
df.columns = df.columns.str.lower()  # Convert to lowercase

# Check if 'Phones' (or its equivalent) exists
expected_col = '_id'  # Adjust based on printed column names
if expected_col not in df.columns:
    raise KeyError(f"Column '{expected_col}' not found in the dataset. Available columns: {df.columns}")

# Split and explode phone numbers
df[expected_col] = df[expected_col].astype(str)
df[expected_col] = df[expected_col].apply(lambda x: [phone.strip() for phone in x.split(',') if phone.strip() != ''])
df_stacked = df.explode(expected_col).reset_index(drop=True)

# Save to output file
output_filepath = r"C:\Users\Kompass\Desktop\Test Files\Stack_Output_KID_NotFound_LinkedInExecutives_20-3-2025.xlsx"
df_stacked.to_excel(output_filepath, index=False)
print('Data saved to ' + output_filepath)
print('Done')
