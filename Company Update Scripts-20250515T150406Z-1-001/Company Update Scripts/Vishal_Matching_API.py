import requests
import json
import pandas as pd
from pprint import pprint
import urllib3

# Disable SSL warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Constants
EXCEL_FILE_PATH = r"D:\Kompass India\Backup Import\MLG Data Import Backup\MLG.xlsx"
API_URL = 'https://87.237.188.124/match/matchBulk'
API_USERNAME = 'kompass_in'
API_PASSWORD = 'F46uEJZ9xAbTMyHT5^$n'
HEADERS = {
    'Content-Type': 'application/json',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
}
SCORE_THRESHOLD = 130  # Threshold for matching score

def read_excel_data(file_path):
    """Read data from an Excel file into a DataFrame."""
    try:
        df = pd.read_excel(file_path)
        print("Excel file read successfully.")
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

import pandas as pd

def create_matching_data_list(df):
    """Create a list of dictionaries from the DataFrame, handling all fields as optional."""
    matching_data_list = []

    for index, row in df.iterrows():
        row_data = {}

        # Handle 'names' field (previously mandatory, now optional)
        try:
            if 'names' in df.columns and not pd.isna(row['names']):
                row_data["name"] = str(row['names']).strip()
            else:
                row_data["name"] = ""  # Default value if 'names' is missing
        except Exception as e:
            print(f"Error processing 'names' in row {index + 1}: {e}")
            row_data["name"] = ""  # Default value if an error occurs

        # Handle 'Company Website' field (optional)
        try:
            if 'Company Website' in df.columns and not pd.isna(row['Company Website']):
                row_data["url"] = str(row['Company Website']).strip()
        except Exception as e:
            print(f"Error processing 'Company Website' in row {index + 1}: {e}")

        # Handle 'Postal Code' field (optional)
        try:
            if 'Postal Code' in df.columns and not pd.isna(row['Postal Code']):
                row_data["zip"] = str(row['Postal Code']).strip()
        except Exception as e:
            print(f"Error processing 'Postal Code' in row {index + 1}: {e}")

        # Handle 'City' field (optional)
        try:
            if 'City' in df.columns and not pd.isna(row['City']):
                row_data["town"] = str(row['City']).strip()
        except Exception as e:
            print(f"Error processing 'City' in row {index + 1}: {e}")

        # Handle 'Address Line 1' field (optional)
        try:
            if 'Address Line 1' in df.columns and not pd.isna(row['Address Line 1']):
                row_data["street"] = str(row['Address Line 1']).strip()
        except Exception as e:
            print(f"Error processing 'Address Line 1' in row {index + 1}: {e}")

        # Handle 'Address Line 2' field (optional)
        try:
            if 'Address Line 2' in df.columns and not pd.isna(row['Address Line 2']):
                row_data["street2"] = str(row['Address Line 2']).strip()
        except Exception as e:
            print(f"Error processing 'Address Line 2' in row {index + 1}: {e}")

        # Handle 'Country' field (previously mandatory, now optional)
        try:
            if 'Country' in df.columns and not pd.isna(row['Country']):
                row_data["country"] = str(row['Country']).strip()
            else:
                row_data["country"] = ""  # Default value if 'Country' is missing
        except Exception as e:
            print(f"Error processing 'Country' in row {index + 1}: {e}")
            row_data["country"] = ""  # Default value if an error occurs

        matching_data_list.append(row_data)

    print(f"Created {len(matching_data_list)} entries for matching.")
    return matching_data_list



def send_api_request(url, username, password, payload, headers):
    """Send a POST request to the API and return the response."""
    try:
        print("Sending payload to API:")
        pprint(payload)  # Debugging: Print the payload being sent
        response = requests.post(url, data=payload, auth=(username, password), verify=False, headers=headers)
        print(response.text)
        response.raise_for_status()  # Raise an exception for HTTP errors
        print("API request successful.")
        return response
    except requests.exceptions.RequestException as e:
        print(f"Error during API request: {e}")
        if hasattr(e, 'response') and e.response:
            print(f"API Response: {e.response.text}")  # Print the API response for debugging
        return None

def process_api_response(response):
    """Process the API response and extract KIDs based on the score threshold."""
    try:
        matching_result_json = json.loads(response.text)  
        kid_list = []

        for item in matching_result_json:
            if 'results' in item and item['results']:
                for result in item['results']:
                    score = result['result_data']['score']
                    kid = result['result_data']['kid']
                    if score >= SCORE_THRESHOLD:
                        kid_list.append(kid)
                    else:
                        kid_list.append('Kid not found')
            else:
                kid_list.append('Kid not found')

        print(f"Processed {len(kid_list)} results from the API.")
        return kid_list
    except Exception as e:
        print(f"Error processing API response: {e}")
        return None

def save_results_to_excel(df, kid_list, file_path):
    """Add the KID list as a new column and save the DataFrame to Excel."""
    df['_id'] = kid_list
    try:
        df.to_excel(file_path, index=False)
        print(f"Results saved to {file_path}.")
    except Exception as e:
        print(f"Error saving results to Excel: {e}")

def main():
    # Step 1: Read the Excel file
    df = read_excel_data(EXCEL_FILE_PATH)
    if df is None:
        return

    # Step 2: Create the matching data list
    matching_data_list = create_matching_data_list(df)
    if matching_data_list is None:
        return

    # Step 3: Create the JSON payload
    result_json = {
        "matching_options": {
            "max_results": 1,
            "include_suppressed": True,
            "api_account": "MyHT95^$nF4EJZ9xAbTn",
            "download": "none",
            "boost_name": 1.50
        },
        "matching_dataList": matching_data_list
    }
    result_json_str = json.dumps(result_json, indent=2)

    # Step 4: Send the API request
    response = send_api_request(API_URL, API_USERNAME, API_PASSWORD, result_json_str, HEADERS)
    if response is None:
        return

    # Step 5: Process the API response
    kid_list = process_api_response(response)
    if kid_list is None:
        return

    # Step 6: Save the results to Excel
    save_results_to_excel(df, kid_list, EXCEL_FILE_PATH)

if __name__ == "__main__":
    main()