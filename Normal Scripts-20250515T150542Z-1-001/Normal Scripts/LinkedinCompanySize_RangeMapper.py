
import pandas as pd

# Define the updated classification mapping with lowercase keys for case-insensitivity
classification_mapping = {
    'self-employed': '0-9',
    '1-10 employees': '0-9',
    '2-10 employees': '0-9',
    '11-50 employees': '20-49',
    '51-200 employees': '100-249',
    '201-500 employees': '250-499',
    '501-1,000 employees': '500-999',
    '1,001-5,000 employees': '1000-4999',
    '5,001-10,000 employees': 'More than 5000',
    '10,000-50,000 employees': 'More than 5000',
    '10,001+ employees': 'More than 5000',
}

def classify_company_size(input_file, output_file, sheet_name=0):
    # Read the Excel file
    try:
        df = pd.read_excel(input_file, sheet_name=sheet_name)
        print(f"Successfully read the sheet. Columns: {df.columns}")
    except Exception as e:
        print(f"Error reading the Excel file: {e}")
        return

    # Check if 'companySize' column exists
    if 'companySize' not in df.columns:
        raise ValueError("The input file must contain a 'companySize' column.")
    
    # Ensure all values are treated as strings and safely apply classification
    def safe_map(x):
        # Convert to string if not NaN, then strip whitespace and map
        if isinstance(x, str):
            return classification_mapping.get(x.strip().lower(), None)
        return None

    # Apply safe mapping to the 'companySize' column
    df['ClassifiedRange'] = df['companySize'].map(safe_map)

    # Save the updated file
    try:
        df.to_excel(output_file, index=False)
        print(f"Classification completed! Updated file saved as {output_file}")
    except Exception as e:
        print(f"Error saving the file: {e}")

# Example usage
input_file = r"C:\Users\Kompass\Downloads\Output_DHL INDIA_Database Sourcing VIA Company List_DataSheets_KOMPASS_10.04.2025.xlsx"
output_file = r"C:\Users\Kompass\Downloads\Output_DHL2 INDIA_Database Sourcing VIA Company List_DataSheets_KOMPASS_10.04.2025.xlsx"

# Explicitly specify the sheet index as 0 for the first sheet
classify_company_size(input_file, output_file, sheet_name=0)


