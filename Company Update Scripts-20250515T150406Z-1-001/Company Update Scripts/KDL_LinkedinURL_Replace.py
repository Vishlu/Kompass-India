import pandas as pd
import json
import requests
import os
import datetime
from tqdm import tqdm
from urllib.parse import urlparse

# Get current date
current_date = datetime.datetime.now()
print('Current date is:', current_date)

# Format the date components
formatted_date = current_date.strftime("KDL_%d_%m_%Y")
source_variable = str(formatted_date).strip()

# Social media keywords
SOCIAL_MEDIA_KEYWORDS = [
    "facebook.com", "twitter.com", "linkedin.com",
    "instagram.com", "youtube.com", "meta.com", "x.com"
]

def is_social_media_url(url):
    """Check if URL is a social media link"""
    if pd.isna(url):
        return False
    return any(keyword in url.lower() for keyword in SOCIAL_MEDIA_KEYWORDS)

def is_linkedin_url(url):
    """Check if URL is LinkedIn"""
    if pd.isna(url):
        return False
    return "linkedin.com" in url.lower()

def get_domain(url):
    """Extract domain from URL"""
    try:
        netloc = urlparse(url).netloc
        if netloc.startswith('www.'):
            netloc = netloc[4:]
        return netloc.lower()
    except:
        return None

def read_excel(file_path):
    """Read Excel file and clean data"""
    df = pd.read_excel(file_path)
    df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
    df.columns = [col.strip() if isinstance(col, str) else col for col in df.columns]
    return df

def get_comp_json(company_id, session_cookie):
    """Fetch company data from Kompass API"""
    base_url = "https://translations.kompass-int.com/kompass/companies/"
    url = base_url + company_id.strip('"')
    
    headers = {
        'User-Agent': 'Mozilla/5.0',
        'Cookie': session_cookie
    }

    try:
        response = requests.get(url, auth=('kin', 'U&8#JjnUMzR&Yj'), headers=headers)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f'Error fetching data for company {company_id}: {e}')
        return None

def write_to_json(json_data, file_path):
    """Write data to JSON file"""
    with open(file_path, 'w', encoding='utf-8') as json_file:
        json.dump(json_data, json_file, indent=4)

def post_comp_json(company_id, updated_json, session_cookie):
    """Post updated data back to Kompass"""
    base_url = "https://translations.kompass-int.com/kompass/companies/"
    url = base_url + company_id.strip('"')

    headers = {
        'User-Agent': 'Mozilla/5.0',
        'Cookie': session_cookie,
        'Content-Type': 'application/json'
    }

    try:
        response = requests.post(url, json=updated_json, auth=('kin', 'U&8#JjnUMzR&Yj'), headers=headers)
        if response.status_code == 200:
            print(f"Company {company_id} updated successfully.")
        else:
            print(f"Failed to update company {company_id}. Status: {response.status_code}")
    except requests.exceptions.RequestException as e:
        print(f"Error updating company {company_id}: {e}")

def process_urls(url_string):
    """Extract official website and LinkedIn URL from input string"""
    if pd.isna(url_string):
        return None, None
    
    urls = [url.strip() for url in str(url_string).split(",") if url.strip()]
    official = None
    linkedin = None
    
    for url in urls:
        if is_linkedin_url(url):
            linkedin = url
        elif not official and not is_social_media_url(url):
            official = url
    
    return official, linkedin

def update_company_websites(website_data, company_json):
    """Update company websites with proper ordering and cleanup"""
    try:
        if 'webSites' not in company_json:
            company_json['webSites'] = []
        
        # Process input URLs
        url_string = website_data['webSites'].iloc[0]
        official_url, linkedin_url = process_urls(url_string)
        
        # Remove all LinkedIn URLs if we have new ones
        if linkedin_url:
            company_json['webSites'] = [
                w for w in company_json['webSites'] 
                if not is_linkedin_url(w['url'])
            ]
        
        # Remove old official website if we have new one
        if official_url:
            # Get domain of new official website
            new_domain = get_domain(official_url)
            if new_domain:
                # Remove any existing URLs with same domain
                company_json['webSites'] = [
                    w for w in company_json['webSites'] 
                    if get_domain(w['url']) != new_domain or is_social_media_url(w['url'])
                ]
        
        # Add new official website (first position)
        if official_url:
            company_json['webSites'].insert(0, {
                'url': official_url,
                'source': source_variable
            })
        
        # Add new LinkedIn URL (second position if official exists, first otherwise)
        if linkedin_url:
            insert_pos = 1 if official_url else 0
            company_json['webSites'].insert(insert_pos, {
                'url': linkedin_url,
                'source': source_variable
            })
            
    except Exception as e:
        print(f'Error updating websites: {e}')

def main():
    # Configuration
    excel_file_path = r"C:\Users\Kompass\Desktop\Website Upload(Official Website & Company LinkedIn)\TestWebsites_25-03-2025.xlsx"  # Update this path
    output_directory = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
    os.makedirs(output_directory, exist_ok=True)
    output_json_file = os.path.join(output_directory, "output.json")
    session_cookie = r"connect.sid=s%3Axinpmg7ACxfaYPqBVTxU-ZlO2JiGp7QH.wkaZ8GgvOTVBlafB9A8ITFFfoOEogQ8rXEWEqr9JtC0"

    print('Output JSON path:', output_json_file)
    
    # Read input file
    excel_file = read_excel(excel_file_path)
    if excel_file is None or excel_file.empty:
        print("Failed to read the Excel file or file is empty, exiting.")
        return

    print('Excel File Sample:')
    print(excel_file.head())
    
    # Group by company ID
    grouped_data = excel_file.groupby('_id')
    
    # Process each company
    for company_id, group in tqdm(grouped_data, desc='Processing Companies'):
        print(f"\nProcessing company ID: {company_id}")
        
        # Get current company data
        company_json = get_comp_json(company_id, session_cookie)
        if not company_json:
            continue
        
        # Update websites
        update_company_websites(group, company_json)
        
        # Post updates to Kompass DB
        post_comp_json(company_id, company_json, session_cookie)
        
        # Save to output JSON
        write_to_json(company_json, output_json_file)

if __name__ == "__main__":
    main()