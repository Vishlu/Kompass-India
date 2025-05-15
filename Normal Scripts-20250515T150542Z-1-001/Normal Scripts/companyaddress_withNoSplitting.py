"""
This is the final scrip which collect only company address

"""


import google.generativeai as genai
import pandas as pd
from tqdm import tqdm
import time
import random

# Configure API Key
API_KEY = "AIzaSyAFZVvhp1DmElGer7WbwoYCehnFj8VRElQ"  # Replace with your actual API key
genai.configure(api_key=API_KEY)
model = genai.GenerativeModel("gemini-2.0-flash")

def get_company_address(company_name, website):
    """Get the complete Indian head office address in raw format"""
    prompt = f"""

    Help me collect the Indian head office address for company {company_name} {website}.
    give me a formatted address in the following format:
    (only collect the head office or corporate office address that is based in India and
    do not provide any other text apart from the address in the specified format).

    Address line 1
    Address lne 2
    City
    State
    Pin code
    Country

    """

    try:
        # Random delay to avoid rate limiting (1-3 seconds)
        time.sleep(random.uniform(1, 3))

        response = model.generate_content(prompt)
        return response.text.strip()
    except Exception as e:
        if "429" in str(e):
            time.sleep(10)  # Longer delay if rate-limited
            return get_company_address(company_name, website)  # Retry once
        return f"Error: {str(e)}"

def process_excel(input_file, output_file):
    """Process Excel file and store full address in one column"""
    df = pd.read_excel(input_file)

    # Validate required columns
    for col in ['Company Name', 'Website']:
        if col not in df.columns:
            raise ValueError(f"Excel file must contain '{col}' column")

    # Add/keep a single column for the full address
    if 'Full Address' not in df.columns:
        df['Full Address'] = ""

    # Process each company with progress bar
    for index, row in tqdm(df.iterrows(), total=len(df), desc="Fetching Addresses"):
        if pd.notna(row['Website']) and pd.notna(row['Company Name']):
            address = get_company_address(row['Company Name'], row['Website'])
            df.at[index, 'Full Address'] = address

    # Save results
    df.to_excel(output_file, index=False)
    print(f"\n✅ Success! Processed {len(df)} companies")
    print(f"📁 Output saved to: {output_file}")

if __name__ == "__main__":
    input_path = "/content/drive/MyDrive/KOMPASS INDIA UPLOAD/Excelsheet/Sample_Company_Description(activities)_29-03-2025.xlsx"
    output_path = "/content/drive/MyDrive/KOMPASS INDIA UPLOAD/Excelsheet/Sample_Output_Address_02-04-2025.xlsx"
    process_excel(input_path, output_path)