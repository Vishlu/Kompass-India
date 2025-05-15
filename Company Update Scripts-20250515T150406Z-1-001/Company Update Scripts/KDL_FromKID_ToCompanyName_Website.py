import requests
import json
import pandas as pd
from tqdm import tqdm
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urlparse

# Elasticsearch URL and credentials
ELASTICSEARCH_URL = "https://translations.kompass-int.com/ldb_in/companies/_search"
AUTH = ("kin", "U&8#JjnUMzR&Yj")

# Function to check if URL is social media
def is_social_media(url):
    if not url or not isinstance(url, str):
        return False
    try:
        domain = urlparse(url.lower()).netloc
        return any(social in domain for social in ['linkedin.com', 'facebook.com'])
    except:
        return False

# Function to query Elasticsearch for a given companyId
def query_elasticsearch(company_id):
    try:
        request_body = {
            "query": {
                "term": {
                    "_id": company_id
                }
            },
            "size": 1
        }

        response = requests.post(
            ELASTICSEARCH_URL,
            json=request_body,
            headers={
                "Content-Type": "application/json",
                "Accept": "application/json",
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.138 Safari/537.36"
            },
            auth=AUTH
        )
        response.raise_for_status()
        return response.json()

    except requests.exceptions.RequestException as e:
        print(f"Error querying Elasticsearch for {company_id}: {e}")
        return None

# Function to process a single row
def process_row(row):
    company_id = row['_id'].strip() if pd.notna(row['_id']) else ""

    if not company_id:
        print(f"Skipping row: Empty or invalid _id.")
        return {**row.to_dict(), 'CompanyName': 'Invalid _id', 'Website': 'Invalid _id'}

    print(f"Querying Elasticsearch for: {company_id}")
    result = query_elasticsearch(company_id)

    if result and 'hits' in result and 'hits' in result['hits'] and result['hits']['hits']:
        source_data = result['hits']['hits'][0]['_source']
        
        # Get company name (first name in names array)
        company_name = source_data.get('names', [{}])[0].get('name', 'N/A')
        
        # Get website (skip social media)
        websites = source_data.get('webSites', [])
        website = 'N/A'
        
        # Iterate through websites to find first non-social media URL
        for site in websites:
            url = site.get('url', '')
            if url and not is_social_media(url):
                website = url
                break
        
        # Create a dictionary with all original columns and the new fields
        result_row = row.to_dict()
        result_row['CompanyName'] = company_name
        result_row['Website'] = website
        return result_row
    else:
        print(f"No results found for: {company_id}")
        return {**row.to_dict(), 'CompanyName': 'Company not found', 'Website': 'Company not found'}

# Function to process input file and write results to output file
def process_file(input_file, output_file):
    try:
        df = pd.read_excel(input_file)

        if '_id' not in df.columns:
            print("Error: Input file must contain a '_id' column.")
            return

        results = []

        with ThreadPoolExecutor(max_workers=10) as executor:
            futures = [executor.submit(process_row, row) for _, row in df.iterrows()]

            for future in tqdm(as_completed(futures), total=len(futures), desc="Processing rows"):
                result_row = future.result()
                if result_row:
                    results.append(result_row)

        output_df = pd.DataFrame(results)

        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            output_df.to_excel(writer, index=False)
        print(f"Results written to {output_file}")

    except Exception as e:
        print(f"Error processing file: {e}")

if __name__ == '__main__':
    input_file = r"C:\Users\Kompass\Downloads\Anjali_Products_Description.xlsx"
    output_file = r"C:\Users\Kompass\Downloads\Anjali_Products_Description.xlsx"

    process_file(input_file, output_file)