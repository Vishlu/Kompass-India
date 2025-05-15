"""
This is the final script for which splits Full Address column into structured components using Gemini API
"""

import google.generativeai as genai
import pandas as pd
from tqdm import tqdm
import time
import random

# Configure API Key
API_KEY = "AIzaSyAFZVvhp1DmElGer7WbwoYCehnFj8VRElQ"  # Replace with your actual API key
genai.configure(api_key=API_KEY)
model = genai.GenerativeModel("gemini-2.0-flash")  # Using flash model for better rate limits

def split_address(full_address, retry_count=0):
    """Split address into components using Gemini API with enhanced error handling"""
    if pd.isna(full_address) or not str(full_address).strip():
        return ["Not available"] * 6

    prompt = f"""
    Split this address into components exactly in this format (return ONLY the 6 lines):
    Address line 1
    Address line 2
    City
    State
    Pin code
    Country

    Input address: "{full_address}"

    Rules:
    1. Fill "Not available" for missing components
    2. Keep original capitalization
    3. Never add explanations
    4. If input is clearly not an address, return "Not available" for all lines
    """

    try:
        # Random delay to avoid rate limiting (1-3 seconds)
        time.sleep(random.uniform(1, 3))

        response = model.generate_content(prompt)
        lines = [line.strip() for line in response.text.split('\n') if line.strip()]

        # Ensure we always return 6 components
        return lines[:6] + ["Not available"] * (6 - len(lines)) if lines else ["Not available"] * 6

    except Exception as e:
        if "429" in str(e):
            if retry_count < 2:  # Maximum 2 retries
                print(f"Rate limited, retrying ({retry_count + 1}/2)...")
                time.sleep(10)  # Longer delay if rate-limited
                return split_address(full_address, retry_count + 1)
            return ["Error: Rate limit exceeded"] * 6
        print(f"Error processing address: {str(e)}")
        return ["Error"] * 6

def process_excel(input_file, output_file):
    """Process Excel file to split addresses"""
    df = pd.read_excel(input_file)

    # Validate required column
    if 'Full Address' not in df.columns:
        raise ValueError("Excel file must contain 'Full Address' column")

    # Define output columns
    address_components = [
        'Address line 1',
        'Address line 2',
        'City',
        'State',
        'Pin code',
        'Country'
    ]

    # Add new columns if they don't exist
    for col in address_components:
        if col not in df.columns:
            df[col] = ""

    # Process each address with progress bar
    for index, row in tqdm(df.iterrows(), total=len(df), desc="Splitting Addresses"):
        components = split_address(row['Full Address'])
        for i, col in enumerate(address_components):
            df.at[index, col] = components[i]

    # Save results
    df.to_excel(output_file, index=False)
    print(f"\n✅ Success! Processed {len(df)} addresses")
    print(f"📁 Output saved to: {output_file}")

if __name__ == "__main__":
    input_path = "/content/drive/MyDrive/KOMPASS INDIA UPLOAD/Excelsheet/Sample_Output_Address_02-04-2025.xlsx"
    output_path = "/content/drive/MyDrive/KOMPASS INDIA UPLOAD/Excelsheet/Split_Sample_Output_Address_02-04-2025.xlsx"
    process_excel(input_path, output_path)