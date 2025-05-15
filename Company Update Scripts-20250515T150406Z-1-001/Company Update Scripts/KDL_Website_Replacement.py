
#final code 

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

# Format the date components as required (KDL_day_month_year)
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

def is_facebook_url(url):
    """Check if URL is Facebook"""
    if pd.isna(url):
        return False
    return "facebook.com" in url.lower()

def is_instagram_url(url):
    """Check if URL is Instagram"""
    if pd.isna(url):
        return False
    return "instagram.com" in url.lower()

def is_twitter_url(url):
    """Check if URL is Twitter"""
    if pd.isna(url):
        return False
    return "twitter.com" in url.lower() or "x.com" in url.lower()

def is_youtube_url(url):
    """Check if URL is YouTube"""
    if pd.isna(url):
        return False
    return "youtube.com" in url.lower()

def normalize_official_url(url):
    """Convert official URL to clean, standardized format:
       1. Remove protocol (http/https)
       2. Ensure www prefix
       3. Remove any path/query parameters
       4. Convert to lowercase
    """
    if pd.isna(url) or not isinstance(url, str):
        return url
    
    try:
        # Remove protocol and path
        parsed = urlparse(url)
        domain = parsed.netloc if parsed.netloc else parsed.path.split('/')[0]
        
        # Ensure consistent www prefix
        if not domain.startswith('www.'):
            domain = 'www.' + domain
        
        # Remove any port numbers if present
        domain = domain.split(':')[0]
        
        # Return in consistent format (www.domain.com)
        return domain.lower()
    except:
        return url.lower()
    

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
    """Extract official website and social media URLs from input string"""
    if pd.isna(url_string):
        return None, None, None, None, None, None
    
    urls = [url.strip() for url in str(url_string).split(",") if url.strip()]
    official = None
    linkedin = None
    facebook = None
    instagram = None
    twitter = None
    youtube = None
    
    for url in urls:
        if is_linkedin_url(url):
            linkedin = url
        elif is_facebook_url(url):
            facebook = url
        elif is_instagram_url(url):
            instagram = url
        elif is_twitter_url(url):
            twitter = url
        elif is_youtube_url(url):
            youtube = url
        elif not official and not is_social_media_url(url):
            official = url
    
    return official, linkedin, facebook, instagram, twitter, youtube

def is_twitter_url(url):
    """Check if URL is Twitter (including x.com)"""
    if pd.isna(url):
        return False
    return "twitter.com" in url.lower() or "x.com" in url.lower()

def update_company_websites(website_data, company_json):
    """Update company websites with exact domain matching and replacement"""
    try:
        if 'webSites' not in company_json:
            company_json['webSites'] = []
        
        # Process input URLs
        url_string = website_data['webSites'].iloc[0]
        official_url, linkedin_url, facebook_url, instagram_url, twitter_url, youtube_url = process_urls(url_string)
        
        # Normalize official URL (if provided)
        normalized_official = normalize_official_url(official_url) if official_url else None
        
        # EXTRACT AND CLASSIFY EXISTING URLs
        existing_official = []
        existing_linkedins = []
        existing_facebooks = []
        existing_instagrams = []
        existing_twitters = []
        existing_youtubes = []
        existing_other_social = []
        
        for site in company_json['webSites']:
            if is_linkedin_url(site['url']):
                existing_linkedins.append(site)
            elif is_facebook_url(site['url']):
                existing_facebooks.append(site)
            elif is_instagram_url(site['url']):
                existing_instagrams.append(site)
            elif is_twitter_url(site['url']):
                existing_twitters.append(site)
            elif is_youtube_url(site['url']):
                existing_youtubes.append(site)
            elif is_social_media_url(site['url']):
                existing_other_social.append(site)
            else:
                # Normalize existing official URLs for comparison
                existing_normalized = normalize_official_url(site['url'])
                existing_official.append((existing_normalized, site))
        
        # 1. HANDLE OFFICIAL WEBSITE (UPDATED LOGIC)
        final_official = []
        if normalized_official:
            new_domain = normalized_official.replace('www.', '').lower()
            matching_existing = []
            non_matching_existing = []
            
            # Separate existing official URLs into matching and non-matching
            for norm_url, site in existing_official:
                existing_domain = norm_url.replace('www.', '').lower()
                if existing_domain == new_domain:
                    matching_existing.append(site)
                else:
                    non_matching_existing.append(site)
            
            # Only add new URL if it matches an existing one
            if matching_existing:
                final_official = [{
                    'url': f"http://{normalized_official}",
                    'source': source_variable
                }]
            else:
                # Keep all existing official URLs if no match
                final_official = [site for _, site in existing_official]
        else:
            # Keep existing official URLs if no new one provided
            final_official = [site for _, site in existing_official]
        
        # 2. HANDLE SOCIAL MEDIA (convert to http) - KEEP EXISTING CODE
        def ensure_http(url):
            if pd.isna(url):
                return url
            if url.startswith('https://'):
                return url.replace('https://', 'http://', 1)
            elif not url.startswith('http://'):
                return 'http://' + url
            return url
        
        # Process LinkedIn
        if linkedin_url:
            final_linkedins = [{
                'url': ensure_http(linkedin_url),
                'source': source_variable
            }]
        else:
            final_linkedins = existing_linkedins
        
        # Process Facebook
        if facebook_url:
            final_facebooks = [{
                'url': ensure_http(facebook_url),
                'source': source_variable
            }]
        else:
            final_facebooks = existing_facebooks
        
        # Process Instagram
        if instagram_url:
            final_instagrams = [{
                'url': ensure_http(instagram_url),
                'source': source_variable
            }]
        else:
            final_instagrams = existing_instagrams
        
        # Process Twitter/X
        if twitter_url:
            final_twitters = [{
                'url': ensure_http(twitter_url),
                'source': source_variable
            }]
        else:
            final_twitters = existing_twitters
        
        # Process YouTube
        if youtube_url:
            final_youtubes = [{
                'url': ensure_http(youtube_url),
                'source': source_variable
            }]
        else:
            final_youtubes = existing_youtubes
        
        # Process other social media
        final_other_social = []
        for site in existing_other_social:
            site_copy = site.copy()
            site_copy['url'] = ensure_http(site['url'])
            final_other_social.append(site_copy)
        
        # BUILD FINAL ORDERED LIST
        company_json['webSites'] = (
            final_official +  # Official website (0 or 1 item)
            final_linkedins +  # LinkedIn URLs
            final_facebooks +  # Facebook URLs
            final_instagrams +  # Instagram URLs
            final_twitters +   # Twitter URLs
            final_youtubes +   # YouTube URLs
            final_other_social # Other social media
        )
        
    except Exception as e:
        print(f'Error updating websites: {e}')
    return company_json

def main():
    # Configuration
    excel_file_path = r"C:\Users\Kompass\Desktop\Test Files\Test_Website_9-4-2025.xlsx"
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
        
        # Update websites - pass the group DataFrame
        updated_json = update_company_websites(group, company_json)
        
        # Post updates to Kompass DB
        post_comp_json(company_id, updated_json, session_cookie)
        
        # Save to output JSON
        write_to_json(updated_json, output_json_file)

if __name__ == "__main__":
    main()














# official website and linkedin replace script

# import pandas as pd
# import json
# import requests
# import os
# import datetime
# from tqdm import tqdm
# from urllib.parse import urlparse

# # Get current date
# current_date = datetime.datetime.now()
# print('Current date is:', current_date)

# # Format the date components as required (KDL_day_month_year)
# formatted_date = current_date.strftime("KDL_%d_%m_%Y")
# source_variable = str(formatted_date).strip()

# # Social media keywords
# SOCIAL_MEDIA_KEYWORDS = [
#     "facebook.com", "twitter.com", "linkedin.com", 
#     "instagram.com", "youtube.com", "meta.com", "x.com"
# ]

# def is_social_media_url(url):
#     """Check if URL is a social media link"""
#     if pd.isna(url):
#         return False
#     return any(keyword in url.lower() for keyword in SOCIAL_MEDIA_KEYWORDS)

# def is_linkedin_url(url):
#     """Check if URL is LinkedIn"""
#     if pd.isna(url):
#         return False
#     return "linkedin.com" in url.lower()

# def normalize_official_url(url):
#     """Convert official URL to clean, standardized format:
#        1. Remove protocol (http/https)
#        2. Ensure www prefix
#        3. Remove any path/query parameters
#        4. Convert to lowercase
#     """
#     if pd.isna(url) or not isinstance(url, str):
#         return url
    
#     try:
#         # Remove protocol and path
#         parsed = urlparse(url)
#         domain = parsed.netloc if parsed.netloc else parsed.path.split('/')[0]
        
#         # Ensure consistent www prefix
#         if not domain.startswith('www.'):
#             domain = 'www.' + domain
        
#         # Remove any port numbers if present
#         domain = domain.split(':')[0]
        
#         # Return in consistent format (www.domain.com)
#         return domain.lower()
#     except:
#         return url.lower()
    

# def read_excel(file_path):
#     """Read Excel file and clean data"""
#     df = pd.read_excel(file_path)
#     df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
#     df.columns = [col.strip() if isinstance(col, str) else col for col in df.columns]
#     return df

# def get_comp_json(company_id, session_cookie):
#     """Fetch company data from Kompass API"""
#     base_url = "https://translations.kompass-int.com/kompass/companies/"
#     url = base_url + company_id.strip('"')
    
#     headers = {
#         'User-Agent': 'Mozilla/5.0',
#         'Cookie': session_cookie
#     }

#     try:
#         response = requests.get(url, auth=('kin', 'U&8#JjnUMzR&Yj'), headers=headers)
#         response.raise_for_status()
#         return response.json()
#     except requests.exceptions.RequestException as e:
#         print(f'Error fetching data for company {company_id}: {e}')
#         return None

# def write_to_json(json_data, file_path):
#     """Write data to JSON file"""
#     with open(file_path, 'w', encoding='utf-8') as json_file:
#         json.dump(json_data, json_file, indent=4)

# def post_comp_json(company_id, updated_json, session_cookie):
#     """Post updated data back to Kompass"""
#     base_url = "https://translations.kompass-int.com/kompass/companies/"
#     url = base_url + company_id.strip('"')

#     headers = {
#         'User-Agent': 'Mozilla/5.0',
#         'Cookie': session_cookie,
#         'Content-Type': 'application/json'
#     }

#     try:
#         response = requests.post(url, json=updated_json, auth=('kin', 'U&8#JjnUMzR&Yj'), headers=headers)
#         if response.status_code == 200:
#             print(f"Company {company_id} updated successfully.")
#         else:
#             print(f"Failed to update company {company_id}. Status: {response.status_code}")
#     except requests.exceptions.RequestException as e:
#         print(f"Error updating company {company_id}: {e}")

# def process_urls(url_string):
#     """Extract official website and LinkedIn URL from input string"""
#     if pd.isna(url_string):
#         return None, None
    
#     urls = [url.strip() for url in str(url_string).split(",") if url.strip()]
#     official = None
#     linkedin = None
    
#     for url in urls:
#         if is_linkedin_url(url):
#             linkedin = url
#         elif not official and not is_social_media_url(url):
#             official = url
    
#     return official, linkedin

# def update_company_websites(website_data, company_json):
#     """Update company websites with perfect handling of all cases"""
#     try:
#         if 'webSites' not in company_json:
#             company_json['webSites'] = []
        
#         # Process input URLs
#         url_string = website_data['webSites'].iloc[0]
#         official_url, linkedin_url = process_urls(url_string)
        
        
#          # Normalize official URL (if provided)
#         normalized_official = normalize_official_url(official_url) if official_url else None
#         # EXTRACT AND CLASSIFY EXISTING URLs
#         existing_official = []
#         existing_linkedins = []
#         existing_social = []
        
#         for site in company_json['webSites']:
#             if is_linkedin_url(site['url']):
#                 existing_linkedins.append(site)
#             elif is_social_media_url(site['url']):
#                 existing_social.append(site)
#             else:
#                 # Normalize existing official URLs for comparison
#                 existing_normalized = normalize_official_url(site['url'])
#                 existing_official.append((existing_normalized, site))
        
#         # In the update_company_websites function, replace the official website handling section with:

#         # 1. HANDLE OFFICIAL WEBSITE
#         if normalized_official:
#             # Check if new URL matches any existing URL when normalized
#             should_replace = False
            
#             if existing_official:
#                 # Get the normalized version of the first existing official URL
#                 existing_normalized = existing_official[0][0]
                
#                 # Compare domains (without www) to see if they match
#                 domain1 = existing_normalized.replace('www.', '')
#                 domain2 = normalized_official.replace('www.', '')
                
#                 # If domains match, we'll replace with the new normalized version
#                 should_replace = (domain1 == domain2)
            
#             # Replace if either:
#             # - No existing official URL, or
#             # - Domains match (different protocol/www variations)
#             if not existing_official or should_replace:
#                 final_official = [{
#                     'url': f"http://{normalized_official}",  # Standardize to http
#                     'source': source_variable
#                 }]
#             else:
#                 # Keep existing URL if domains don't match
#                 final_official = [existing_official[0][1]]
#         else:
#             # Keep first existing official URL if available
#             final_official = [existing_official[0][1]] if existing_official else []
        
#         # 2. HANDLE LINKEDIN
#         if linkedin_url:
#             # Use new LinkedIn URL and discard old ones
#             final_linkedins = [{
#                 'url': linkedin_url,
#                 'source': source_variable
#             }]
#         else:
#             # Keep all existing LinkedIn URLs
#             final_linkedins = existing_linkedins
        
#         # 3. PRESERVE SOCIAL MEDIA
#         final_social = existing_social
        
#         # BUILD FINAL ORDERED LIST
#         company_json['webSites'] = (
#             final_official +  # Official website (0 or 1 item)
#             final_linkedins +  # LinkedIn URLs (0-N items)
#             final_social  # Other social media (0-N items)
#         )
        
#     except Exception as e:
#         print(f'Error updating websites: {e}')
#     return company_json

# def main():
#     # Configuration
#     excel_file_path = r"C:\Users\Kompass\Desktop\Website Upload(Official Website & Company LinkedIn)\Websites\Uploaded_Found_2_Website_21-3-2025.xlsx"
#     output_directory = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
#     os.makedirs(output_directory, exist_ok=True)
#     output_json_file = os.path.join(output_directory, "output.json")
#     session_cookie = r"connect.sid=s%3Axinpmg7ACxfaYPqBVTxU-ZlO2JiGp7QH.wkaZ8GgvOTVBlafB9A8ITFFfoOEogQ8rXEWEqr9JtC0"

#     print('Output JSON path:', output_json_file)
    
#     # Read input file
#     excel_file = read_excel(excel_file_path)
#     if excel_file is None or excel_file.empty:
#         print("Failed to read the Excel file or file is empty, exiting.")
#         return

#     print('Excel File Sample:')
#     print(excel_file.head())
    
#     # Group by company ID
#     grouped_data = excel_file.groupby('_id')
    
#     # Process each company
#     for company_id, group in tqdm(grouped_data, desc='Processing Companies'):
#         print(f"\nProcessing company ID: {company_id}")
        
#         # Get current company data
#         company_json = get_comp_json(company_id, session_cookie)
#         if not company_json:
#             continue
        
#         # Update websites - pass the group DataFrame
#         updated_json = update_company_websites(group, company_json)
        
#         # Post updates to Kompass DB
#         post_comp_json(company_id, updated_json, session_cookie)
        
#         # Save to output JSON
#         write_to_json(updated_json, output_json_file)

# if __name__ == "__main__":
#     main()