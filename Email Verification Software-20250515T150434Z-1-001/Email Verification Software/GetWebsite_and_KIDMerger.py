import pandas as pd
import json
import requests
import os
import datetime
from tqdm import tqdm
import tempfile

# Get current date
current_date = datetime.datetime.now()
print('Current date is:', current_date)

# Format the date components as required (KDL_day_month_year)
formatted_date = current_date.strftime("KDL_%d_%m_%Y")
source_variable = str(formatted_date).strip()

def read_excel(file_path):
    try:
        # Read Excel file into DataFrame
        df = pd.read_excel(file_path)
        # Apply strip() function to each cell value in DataFrame
        df = df.apply(lambda x: x.strip() if isinstance(x, str) else x)
        df.columns = [col.strip() if isinstance(col, str) else col for col in df.columns]
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

def get_comp_json(company_id, session_cookie):
    base_url = "https://translations.kompass-int.com/kompass/companies/"
    url = base_url + company_id.strip('"')
    session = requests.Session()
    session.headers.update({
        "Cookie": session_cookie,
        'User-Agent': 'Mozilla/5.0'
    })
    
    try:
        response = session.get(url, auth=('kin', 'U&8#JjnUMzR&Yj'))
        response.raise_for_status()  # Will raise an error for 4xx or 5xx status codes
        return response
    except requests.exceptions.RequestException as e:
        print(f"Error fetching data for company {company_id}: {e}")
        return None

def write_to_excel(data, file_path):
    try:
        # Convert data to DataFrame and write to Excel
        df = pd.DataFrame(data)
        df.to_excel(file_path, index=False)
        print(f"Data written to Excel file: {file_path}")
    except Exception as e:
        print(f"Error writing to Excel file: {e}")

def merge_excel_files(file1_path, file2_path, output_file_path):
    try:
        # Read the first Excel file
        df1 = pd.read_excel(file1_path)
        # Read the second Excel file
        df2 = pd.read_excel(file2_path)
        
        # Merge the two dataframes on '_id' column
        merged_df = pd.merge(df1, df2, on='_id', how='inner')
        
        # Save the merged dataframe to a new Excel file
        merged_df.to_excel(output_file_path, index=False)
        print(f"Merged file has been saved at: {output_file_path}")
    except Exception as e:
        print(f"Error occurred: {e}")

def main():
    # Paths and session details
    input_excel_path = r"C:\Users\Kompass\Desktop\Linkedin Unit Manager Email Address\Found KID Data Upload_Unit Manager_SERP Data Extraction_29-01-2025.xlsx"
    session_cookie = r"connect.sid=s%3Axinpmg7ACxfaYPqBVTxU-ZlO2JiGp7QH.wkaZ8GgvOTVBlafB9A8ITFFfoOEogQ8rXEWEqr9JtC0"
    
    # Temporary file handling
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as temp_file1, tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as temp_file2:
        temp_file1_path = temp_file1.name
        temp_file2_path = temp_file2.name

        # First script process
        all_data = []
        input_data = read_excel(input_excel_path)
        if input_data is not None:
            if '_id' not in input_data.columns:
                print("Error: '_id' column not found in Excel file.")
                return

            total_rows = len(input_data)
            for index, row in tqdm(input_data.iterrows(), desc='Progress Bar : ', total=total_rows):
                company_id = row['_id']

                try:
                    # Fetch company JSON
                    response = get_comp_json(company_id, session_cookie)
                    if response and response.status_code == 200:
                        json_data = response.json()
                        websites = json_data.get('webSites', [])
                        if websites:
                            for website in websites:
                                required_data = {
                                    "_id": company_id,
                                    "webSites": website.get('url', 'N/A')
                                }
                                all_data.append(required_data)
                        else:
                            required_data = {"_id": company_id, "webSites": 'N/A'}
                            all_data.append(required_data)
                    else:
                        required_data = {"_id": company_id, "webSites": 'N/A'}
                        all_data.append(required_data)
                except Exception as e:
                    print(f"Error processing company {company_id}: {e}")
                    required_data = {"_id": company_id, "webSites": 'N/A'}
                    all_data.append(required_data)

        # Write first script's output to a temporary file
        write_to_excel(all_data, temp_file1_path)

        # Second script process (merging)
        output_file_path = r"C:\Users\Kompass\Desktop\Linkedin Unit Manager Email Address\Output_Website_Found KID Data Upload_Unit Manager_SERP Data Extraction_29-01-2025.xlsx"
        merge_excel_files(temp_file1_path, input_excel_path, output_file_path)

if __name__ == "__main__":
    main()
