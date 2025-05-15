# Importing the necessary libraries
import pandas as pd
import json
import requests
from pprint import pprint
import os
import math
import datetime
import unicodedata
import re
from tqdm import tqdm
from urllib.parse import urlparse, urlunparse

# Get current date
current_date = datetime.datetime.now()

print('Current date is :',current_date)

# Format the date components as required (KDL_day_month_year)
formatted_date = current_date.strftime("KDL_%d_%m_%Y")

# Assign the formatted date to your source variable
source_variable = str(formatted_date).strip()


# Function to read excel and strip whitespace
def read_excel(file_path):
    # Read Excel file into DataFrame
    df = pd.read_excel(file_path)
    
    # Apply strip() function to each cell value in DataFrame
    df = df.apply(lambda x: x.strip() if isinstance(x, str) else x)

    df.columns = [col.strip() if isinstance(col, str) else col for col in df.columns]

    
    return df


# Function to get Company Json from database
def get_comp_json(company_id, session_cookie):
    base_url = "https://translations.kompass-int.com/kompass/companies/"
    # url = f"{base_url}{company_id.strip('\"')}"
    url = base_url + company_id.strip('"')
    session = requests.Session()

    try:
        session.headers.update({"Cookie": session_cookie})
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.138 Safari/537.36'}
        response = session.get(url, auth=('kin', 'U&8#JjnUMzR&Yj'),headers=headers)
        return response
    except Exception as e:
        print(f'Exception: {e}')
        return None


# Function to write to JSON file
def write_to_json(json_data, file_path):
    with open(file_path, 'w') as json_file:
        json.dump(json_data, json_file, indent=4)


# Function to post the data
def post_comp_json(company_id, output_json_file, session_cookie):
    base_url = "https://translations.kompass-int.com/kompass/companies/"
    # url = f"{base_url}{company_id.strip('\"')}"
    url = base_url + company_id.strip('"')
    session = requests.Session()

    try:
        with open(output_json_file, 'r') as json_file:
            data = json.load(json_file)

        session.headers.update({"Cookie": session_cookie})
        response = session.post(url, json=data, auth=('kin', 'U&8#JjnUMzR&Yj'))

        if response.status_code == 200:
            json_data = response.json()
            print()
            print('Posted data for :', company_id)
            print()
        else:
            print("Unexpected Status Code :", response.status_code)

    except Exception as e:
        print(f'Exception: {e}')
        
        
def custom_title_case(input_string):
    words = input_string.split()
    titlecased_words = [word.title() if word.islower() else word for word in words]
    return ' '.join(titlecased_words)

def update_comp_json(row,company_json,output_json_file):
    # unique_phones = set()
    # unique_area_codes = set()

    try:
        if (
                pd.notna(row.get('salutation')) or pd.notna(row.get('firstName')) or pd.notna(row.get('lastname')) or
                pd.notna(row.get('functionName')) or pd.notna(row.get('executiveEmail')) or pd.notna(row.get('ExecCountryCode')) or
                pd.notna(row.get('ExecMobileNumber')) or pd.notna(row.get('linkedinprofile')) or pd.notna(row.get('functionCode'))
            ):
                executives = company_json.get('executives', [])

                # Determine executive type
                exec_type = 'executive'
                designation = str(row.get('functionName'))
                if designation in ['Chairman', 'Director'] or str(row.get('functionCode')).strip() in ['5101', '5102']:
                    exec_type = 'board'

                # Extract relevant fields
                first_name = (row.get('firstName').split() or [''])[0] if pd.notna(row.get('firstName')) else None
                last_name = (row.get('lastname').split() or [''])[-1] if pd.notna(row.get('lastname')) else None
                first_initial_lastname = last_name[0].upper() if last_name else None  # First initial of last name
                function_code = str(int(row.get('functionCode'))).strip() if pd.notna(row.get('functionCode')) else None
                linkedin_profile = row.get('linkedinprofile').strip() if pd.notna(row.get('linkedinprofile')) else None
                executive_email = row.get('executiveEmail').strip() if pd.notna(row.get('executiveEmail')) else None
                exec_country_code = str(int(row.get('ExecCountryCode'))) if pd.notna(row.get('ExecCountryCode')) else ''
                exec_mobile_number = str(int(row.get('ExecMobileNumber'))) if pd.notna(row.get('ExecMobileNumber')) else ''

                # Flag to check if a match is found
                match_found = False

                # Iterate through existing executives to find a match
                for i, person_info in enumerate(executives):
                    existing_first_name = (person_info.get('firstName', '').split() or [''])[0]
                    existing_last_name = (person_info.get('lastName', '').split() or [''])[-1]
                    existing_first_initial_lastname = existing_last_name[0].upper() if existing_last_name else None
                    existing_social_networks = [entry.get("address", "").strip() for entry in person_info.get("socialNetworks", []) if entry.get("address")]

                    # **Condition 1: If `firstName` and first initial of `lastName` match → Update specific fields**
                    if first_name == existing_first_name and first_initial_lastname == existing_first_initial_lastname:
                        match_found = True
                        # Update specific fields
                        person_info['salutation'] = row.get('salutation') if pd.notna(row.get('salutation')) else person_info.get('salutation')
                        person_info['lastName'] = last_name
                        person_info['mainfunction'] = function_code
                        person_info['functionName'] = row.get('functionName').strip() if pd.notna(row.get('functionName')) else person_info.get('functionName')
                        person_info['executiveType'] = exec_type

                        # Update LinkedIn profile if provided
                        if linkedin_profile:
                            person_info['socialNetworks'] = [{'address': linkedin_profile}]

                        # Update executive email if provided
                        if executive_email:
                            person_info['emailAddress'] = executive_email

                        # Update mobile phone number if provided
                        if exec_country_code and exec_mobile_number:
                            person_info['mobilePhoneNumber'] = {
                                "countryPhoneCode": exec_country_code,
                                "number": exec_mobile_number,
                                "islocal": False
                            }
                        break  # Exit the loop once a match is found and updated

                    # **Condition 2: If `linkedinprofile` matches → Update specific fields**
                    if linkedin_profile in existing_social_networks:
                        match_found = True
                        # Update specific fields
                        person_info['salutation'] = row.get('salutation') if pd.notna(row.get('salutation')) else person_info.get('salutation')
                        person_info['firstName'] = first_name
                        person_info['lastName'] = last_name
                        person_info['mainfunction'] = function_code
                        person_info['functionName'] = row.get('functionName').strip() if pd.notna(row.get('functionName')) else person_info.get('functionName')
                        person_info['executiveType'] = exec_type

                        # Update LinkedIn profile if provided
                        if linkedin_profile:
                            person_info['socialNetworks'] = [{'address': linkedin_profile}]

                        # Update executive email if provided
                        if executive_email:
                            person_info['emailAddress'] = executive_email

                        # Update mobile phone number if provided
                        if exec_country_code and exec_mobile_number:
                            person_info['mobilePhoneNumber'] = {
                                "countryPhoneCode": exec_country_code,
                                "number": exec_mobile_number,
                                "islocal": False
                            }
                        break  # Exit the loop once a match is found and updated

                # **Condition 3: If no match found → Add as a new executive**
                if not match_found and linkedin_profile and all(x is not None for x in [first_name, last_name, function_code]):
                    new_exec_info = {
                        'isoCode': 'en',
                        "executiveType": exec_type,
                        "mainfunction": function_code,
                        "salutation": row.get('salutation') if pd.notna(row.get('salutation')) else '',
                        "firstName": first_name,
                        "lastName": last_name,
                        "functionName": row.get('functionName').strip() if pd.notna(row.get('functionName')) else '',
                        "emailAddress": executive_email,
                        'order': 0,
                        'source': source_variable,
                        'mobilePhoneNumber': {
                            "countryPhoneCode": exec_country_code,
                            "number": exec_mobile_number,
                            "islocal": False
                        },
                        'socialNetworks': [{'address': linkedin_profile}]
                    }
                    executives.append(new_exec_info)

            # **Sorting Executives**
        company_json['executives'] = sorted(executives, key=lambda x: x.get('mainfunction', ''), reverse=False)
        order_counter = 1
        for executive in company_json['executives']:
            executive['order'] = order_counter
            order_counter += 1

        # **Deduplicate Executives**
        seen_combinations = set()
        unique_executives = []
        for executive in company_json['executives']:
            key = (executive.get('firstName', ''), executive.get('lastName', ''), executive.get('mainfunction', ''))
            if key not in seen_combinations:
                unique_executives.append(executive)
                seen_combinations.add(key)
        company_json['executives'] = unique_executives

        # **Set executive type after deduplication**
        for executive in company_json['executives']:
            mainfunction = executive.get('mainfunction', '')
            functionName = executive.get('functionName', '')
            if mainfunction in ['5101', '5102'] or functionName in ['Director', 'Chairman']:
                executive['executiveType'] = 'board'

        # **Save updated JSON file**
        with open(output_json_file, 'w') as json_file:
            json.dump(company_json, json_file, indent=4)

    except Exception as e:
        print(f'Exception: {e}')


    # Logic for company name
        
    try:
        if pd.notna(row.get('names')) or pd.notna(row['continuation']):
            continuation = str(row.get('continuation')).strip()
            #print('Current Continuation :', continuation)
            
            if 'names' in company_json:
                #print('Notna condition value :', pd.notna(row.get('continuation')))
                if pd.notna(row.get('continuation')):
                    #print('Inside Continuation condition')
                    if company_json['names']:  # Check if 'names' list is not empty
                        company_json['names'][0]['continuation'] = continuation
                    else:
                        # Append a new entry with provided continuation if 'names' list is empty
                        company_json['names'].append({'isoCode': 'en', 'name': '', 'continuation': continuation})

                    with open(output_json_file, 'w') as json_file:
                        json.dump(company_json, json_file, indent=4)
                    
                comp_names = [name.get('name').lower() for name in company_json.get('names', [])]  # Extract existing names
                comp_names_first_words = [name.split()[0] for name in comp_names]  # Extract first words of existing names
                #print('Current Names :', comp_names)

                new_name_lower = row.get('names', '').lower()
                if new_name_lower not in comp_names and new_name_lower.split()[0] not in comp_names_first_words:
                    company_json['names'].append({
                        'isoCode': 'en',
                        'name': str(row.get('names')).strip() if pd.notna(row.get('names')) else '',
                        'continuation': continuation if pd.notna(row.get('continuation')) else ''
                    })

                    with open(output_json_file, 'w') as json_file:
                        json.dump(company_json, json_file, indent=4)

                    # If company_json['names'] is not empty and company name is not present at the first entry, add it there
                    if company_json['names'] and company_json['names'][0]['name'] == '':
                        company_json['names'][0]['name'] = str(row.get('names')).strip() if pd.notna(row.get('names')) else ''

                        with open(output_json_file, 'w') as json_file:
                            json.dump(company_json, json_file, indent=4)

            else:
                # Create 'names' key and add a new entry with provided name and continuation
                company_json['names'] = [{'isoCode': 'en', 'name': str(row.get('names')).strip() if pd.notna(row.get('names')) else '',
                                        'continuation': continuation if pd.notna(row.get('continuation')) else ''}]
                with open(output_json_file, 'w') as json_file:
                    json.dump(company_json, json_file, indent=4)

    except Exception as e:
        print(f'Exception: {e}')


    # Logic for websites 
    try:

        # Function to normalize a URL
        def normalize_url(url):
            try:
                parsed_url = urlparse(url)
                # Normalize the scheme and remove trailing slashes
                netloc = parsed_url.netloc.lower()
                path = parsed_url.path.rstrip('/')  # Remove trailing slash
                return urlunparse(('https', netloc, path, '', '', ''))  # Always return https and normalized path
            except Exception:
                return url  # Return original if parsing fails

        # Logic for replacing or adding new websites
        if 'webSites' in company_json:
            websites = company_json.get('webSites', [])
            new_websites = row.get('webSites')

            # Split multiple websites if they are comma-separated
            new_websites_list = [
                url.strip() for url in new_websites.split(',')
                if pd.notna(url) and url.strip()  # Ignore empty or NaN values
            ] if pd.notna(new_websites) else []

            # Normalize and process each new website
            for new_website in new_websites_list:
                normalized_new_website = normalize_url(new_website)

                # 1. Remove duplicates based on normalized URLs
                seen_urls = {}
                for website in websites:
                    existing_url = website.get('url')
                    normalized_existing_url = normalize_url(existing_url)
                    seen_urls[normalized_existing_url] = website  # Keep the latest one if duplicates found

                # 2. Check if the new website matches an existing URL (normalized), replace it
                replaced = False
                for website in websites:
                    existing_url = website.get('url')
                    normalized_existing_url = normalize_url(existing_url)

                    if normalized_existing_url == normalized_new_website:
                        # Replace the existing URL with the new one
                        website['url'] = str(new_website) if pd.notna(new_website) else ''
                        website['source'] = source_variable
                        replaced = True
                        break

                # 3. If no replacement occurred, append the new website
                if not replaced and normalized_new_website:
                    company_json['webSites'].append({'url': str(new_website), 'source': source_variable})

                # 4. Update the 'webSites' to only contain unique URLs (based on normalized URLs)
                unique_websites = list(seen_urls.values())

                # Adding the new website URL if it's unique and not in seen_urls
                if normalized_new_website and not any(normalize_url(website['url']) == normalized_new_website for website in unique_websites):
                    unique_websites.append({'url': str(new_website), 'source': source_variable})

                # Update the company_json with the deduplicated list of websites
                company_json['webSites'] = unique_websites

            # Save the updated JSON
            with open(output_json_file, 'w') as json_file:
                json.dump(company_json, json_file, indent=4)

        else:
            # If 'webSites' field doesn't exist, create it with the new website(s)
            new_websites = row.get('webSites')
            new_websites_list = [
                url.strip() for url in new_websites.split(',')
                if pd.notna(url) and url.strip()  # Ignore empty or NaN values
            ] if pd.notna(new_websites) else []

            company_json['webSites'] = [
                {'url': str(website), 'source': source_variable}
                for website in new_websites_list
            ]

            with open(output_json_file, 'w') as json_file:
                json.dump(company_json, json_file, indent=4)


    except Exception as e:
        print(f'Exception: {e}')

    # Handling 'phones' field
    try:

        if pd.notna(row.get('phones')):
            current_phone = unicodedata.normalize('NFC', str(int(row.get('phones'))).strip())
            if 'phones' in company_json:
                phones = company_json.get('phones', [])
                # Clean and normalize phone numbers for comparison
                phone_nos_list =  [unicodedata.normalize('NFC', ''.join(char for char in str(name.get('number')) if char.isdigit())) for name in phones]

                # Clean and normalize the current phone number for comparison
                #print('Phone Nos List :',phone_nos_list)
                #print('Current Phone :',current_phone)
                if current_phone not in phone_nos_list:
                    company_json['phones'].append({
                        'countryPhoneCode': int(row.get('CountryPhoneCode')) if pd.notna(row.get('CountryPhoneCode')) else '',
                        'areaPhoneCode': str(int(row.get('areaPhoneCode'))) if pd.notna(row.get('areaPhoneCode')) else '',
                        'number': current_phone if pd.notna(row.get('phones')) else '',
                        'source' : source_variable
                    })

                    with open(output_json_file, 'w') as json_file:
                        json.dump(company_json, json_file, indent=4)

                elif current_phone in phone_nos_list:
                    matched_index = phone_nos_list.index(current_phone)
                    if(pd.notna(row.get('CountryPhoneCode'))):
                        company_json['phones'][matched_index]["countryPhoneCode"] = int(row.get('CountryPhoneCode')) if pd.notna(row.get('CountryPhoneCode')) else ''
                        company_json['phones'][matched_index]['source'] = source_variable
                    if(pd.notna(row.get('areaPhoneCode'))):
                        company_json['phones'][matched_index]["areaPhoneCode"] = str(int(row.get('areaPhoneCode'))) if pd.notna(row.get('areaPhoneCode')) else ''
                        company_json['phones'][matched_index]['source'] = source_variable

                    with open(output_json_file, 'w') as json_file:
                        json.dump(company_json, json_file, indent=4)
            else:
                # Initialize the set with existing values if 'phones' key is not present
                # unique_phones = {str(name.get('number')) for name in phones}
                # unique_area_codes = {str(name.get('areaPhoneCode')) for name in phones}

                company_json['phones'] = [{
                    'countryPhoneCode': int(row.get('CountryPhoneCode')) if pd.notna(row.get('CountryPhoneCode')) else '',
                    'areaPhoneCode': str(int(row.get('areaPhoneCode'))) if pd.notna(row.get('areaPhoneCode')) else '',
                    'number': current_phone if pd.notna(row.get('phones')) else '',
                    'source' : source_variable
                }]

                with open(output_json_file, 'w') as json_file:
                    json.dump(company_json, json_file, indent=4)

    except Exception as e:
        print(f'Exception: {e}')

    # Handling Emails
    try:
        

        if(pd.notna(row.get('emails'))):

            if 'emails' in company_json:
                emails = company_json.get('emails',[])
                comp_emails = [name.get('value').strip() for name in emails]
                if str(row.get('emails')) not in comp_emails:
                    company_json['emails'].append({'value': str(row.get('emails')) if pd.notna(row.get('emails')) else '','source':source_variable})
                    with open(output_json_file, 'w') as json_file:
                        json.dump(company_json, json_file, indent=4)

                # Removing Duplicates for Emails
                unique_emails = set()
                indices_to_remove = []
                for index, email in enumerate(company_json['emails']):
                    email_value = email.get('value').strip()
                    if email_value in unique_emails:
                        indices_to_remove.append(index)
                    else:
                        unique_emails.add(email_value)

                for index in sorted(indices_to_remove, reverse=True):
                    company_json['emails'].pop(index)

                with open(output_json_file, 'w') as json_file:
                    json.dump(company_json, json_file, indent=4)

                # Increasing Ranking in Ascending Order for Emails
                email_ranking = 1
                for email in company_json['emails']:
                    email['ranking'] = email_ranking
                    email_ranking += 1

                with open(output_json_file, 'w') as json_file:
                    json.dump(company_json, json_file, indent=4)

            else:
                company_json['emails'] = [{'value': str(row.get('emails')) if pd.notna(row.get('emails')) else '','ranking':1,'source':source_variable}]
                with open(output_json_file, 'w') as json_file:
                    json.dump(company_json, json_file, indent=4)

    except Exception as e:
        print(f'Exception: {e}')

    try:

    # Logic for turnover
        if(pd.notna(row.get('turnover'))):
            turnover_dict = {
                    'INR Under 5 Million': 'A',
                    'INR 5 - 10 Million': 'B',
                    'INR 10 - 50 Million': 'C',
                    'INR 50 - 100 Million': 'D',
                    'INR 100 - 500 Million': 'E',
                    'INR 500 - 1000 Million': 'F',
                    'INR 1000 - 5000 Million': 'G',
                    'INR 5000 - 10000 Million': 'H',
                    'INR 10000 - 50000 Million': 'I',
                    'Over 50000 Million': 'J'
                }
            
            
            if 'financialDatas' in company_json:
                # turnover_dict = {
                #     'INR Under 5 Million': 'A',
                #     'INR 5 - 10 Million': 'B',
                #     'INR 10 - 50 Million': 'C',
                #     'INR 50 - 100 Million': 'D',
                #     'INR 100 - 500 Million': 'E',
                #     'INR 500 - 1000 Million': 'F',
                #     'INR 1000 - 5000 Million': 'G',
                #     'INR 5000 - 10000 Million': 'H',
                #     'INR 10000 - 50000 Million': 'I',
                #     'Over 50000 Million': 'J'
                # }
                excel_turnover = str(row.get('turnover')).strip() if row.get('turnover') else ''
                financial_data = company_json['financialDatas']['FinancialData']
                new_turnover = turnover_dict[excel_turnover] 
                existing_turnover_list = [name.get('mainTurnoverRange') for name in financial_data]
                existing_year_list = [data.get('Year') for data in financial_data]
                today = datetime.date.today()
                current_year = today.year

                # Append
                if current_year not in existing_year_list and new_turnover not in existing_turnover_list:
                    company_json['financialDatas']['FinancialData'].append({
                        'Year': current_year,
                        "Currency": "INR",
                        "endMonth": 12,
                        "duration": 12,
                        "mainTurnoverRange": new_turnover if pd.notna(row.get('turnover')) else '',
                        'source' : source_variable
                    })

                    with open(output_json_file, 'w') as json_file:
                        json.dump(company_json, json_file, indent=4)
                
                # If current year is present then update range
                if current_year in existing_year_list:
                    current_year_index = existing_year_list.index(current_year)
                    company_json['financialDatas']['FinancialData'][current_year_index]['mainTurnoverRange'] = new_turnover if pd.notna(row.get('turnover')) else ''
                    company_json['financialDatas']['FinancialData'][current_year_index]['source'] = source_variable

                    with open(output_json_file, 'w') as json_file:
                        json.dump(company_json, json_file, indent=4)

            else:
                today = datetime.date.today()
                current_year = today.year
                excel_turnover = str(row.get('turnover')) 
                new_turnover = turnover_dict[excel_turnover]
                company_json['financialDatas'] = {
                    'confidential': 0,
                    'FinancialData': [{
                        'Year': current_year,
                        "Currency": "INR",
                        "endMonth": 12,
                        "duration": 12,
                        "mainTurnoverRange": new_turnover if pd.notna(row.get('turnover')) else '',
                        'source' : source_variable
                    }]
                }

                with open(output_json_file, 'w') as json_file:
                    json.dump(company_json, json_file, indent=4)

    except Exception as e:
        print(f'Exception: {e}')

    # To update employee range   
    try:

        if pd.notna(row['companySize']):
            employee_range_dict = {
                    'Not declared': '0',
                    '0-9': '1',
                    '10-19': '2',
                    '20-49': '3',
                    '50-99': '4',
                    '100-249': '5',
                    '250-499': '6',
                    '500-999': '7',
                    '1000-4999': '8',
                    'More than 5000': '9'
                }
            if "employeeInfo" in company_json:
                excel_employee_range = str(row.get('companySize')).strip()  # More than 5000
                address_employee_range = company_json.get("employeeInfo", {}).get("addressEmployeRange", "")
                excel_emp_range = employee_range_dict.get(excel_employee_range)
                # Check if excel_employee_range is not None before accessing the dictionary
                if excel_employee_range is not None:
                    excel_emp_range = employee_range_dict.get(excel_employee_range)

                    if excel_emp_range is not None and excel_emp_range != address_employee_range:
                        company_json["employeeInfo"]["addressEmployeRange"] = excel_emp_range
                        company_json["employeeInfo"]["source"] = source_variable
                        with open(output_json_file, 'w') as json_file:
                            json.dump(company_json, json_file, indent=4)

            else:
                excel_employee_range = str(row.get('companySize')).strip()  # More than 5000
                address_employee_range = company_json.get("employeeInfo", {}).get("addressEmployeRange", "")
                excel_emp_range = employee_range_dict.get(excel_employee_range)
                company_json['employeeInfo'] = {"addressEmployeRange": excel_emp_range , 'source' : source_variable}
                with open(output_json_file, 'w') as json_file:
                    json.dump(company_json, json_file, indent=4)

    except Exception as e:
        print(f'Exception: {e}')
    
    
    # Logic for search indicators
    try:
        if pd.notna(row['searchIndicators']):
            if "searchIndicators" in company_json:
                existing_search_indicators_list = company_json.get('searchIndicators', [])

                # Check if the existing search indicators are stored as a list
                if isinstance(existing_search_indicators_list, list):
                    existing_search_indicators = set(existing_search_indicators_list)
                else:
                    existing_search_indicators = set(existing_search_indicators_list.split(','))

                words_from_excel = row['searchIndicators'].split(',')

                # Convert all words to lowercase for case-insensitive comparison
                existing_search_indicators_lower = {word.strip().lower() for word in existing_search_indicators}

                # Find unique words by subtracting existing search indicators from words from Excel
                unique_words = {word.strip().lower() for word in words_from_excel if word.strip().lower()} - existing_search_indicators_lower

                # Debug information
                # print("Debug Info:")
                # print("Existing Search Indicators (set):", existing_search_indicators)
                # print("Unique Words:", unique_words)

                # Add unique words to the existing set
                existing_search_indicators.update(unique_words)

                # Convert each word in existing_search_indicators set to title case using custom_title_case function
                title_case_search_indicators = [custom_title_case(word) for word in existing_search_indicators]

                # Convert the title case search indicators list to a JSON array
                company_json['searchIndicators'] = title_case_search_indicators

                with open(output_json_file, 'w') as json_file:
                    json.dump(company_json, json_file, indent=4)

            else:
                company_json['searchIndicators'] = [custom_title_case(word.strip()) for word in row['searchIndicators'].split(',') if word.strip()]

                with open(output_json_file, 'w') as json_file:
                    json.dump(company_json, json_file, indent=4)

    except Exception as e:
        print(f'Exception: {e}')




    # To update the address field
    try:
        if (pd.notna(row['street']) or pd.notna(row['complement']) or pd.notna(row['zipCode']) or pd.notna(row['city']) or pd.notna(row['Country'])):
            zip_code_to_update = str(int(row.get('zipCode'))).strip()
            if "address" in company_json:
                existing_zip_codes = [address["zipCode"].strip() for address in company_json.get("address", [])]
                #print('Existing Zip codes :',existing_zip_codes)
                #print('Zip Code to update :',zip_code_to_update)

                if zip_code_to_update in existing_zip_codes:
                    # Find the index of the address with the matching zip code
                    index_to_update = existing_zip_codes.index(zip_code_to_update)
                    #print('Index to update :',index_to_update)
                    # Update the fields
                    company_json['address'][index_to_update]['street'] = str(row.get('street')) if pd.notna(row.get('street')) else ''
                    company_json['address'][index_to_update]['complement'] = str(row.get('complement')) if pd.notna(row.get('complement')) else ''
                    company_json['address'][index_to_update]['zipCode'] = zip_code_to_update
                    company_json['address'][index_to_update]['city'] = str(row.get('city')) if pd.notna(row.get('city')) else ''
                    company_json['address'][index_to_update]['country'] = str(row.get('Country')) if pd.notna(row.get('Country')) else ''
                    company_json['address'][index_to_update]['source'] = source_variable

                    if 'type' not in company_json['address'][index_to_update]:
                        company_json['address'][index_to_update]['type'] = 'VIS'
                        company_json['address'][index_to_update]['source'] = source_variable
                    
                    with open(output_json_file, 'w') as json_file:
                        json.dump(company_json, json_file, indent=4)
                else:
                    # If the zip code is not present, add a new address
                    company_json['address'].append({
                        'isoCode': 'en',
                        "street": str(row.get('street')) if pd.notna(row.get('street')) else '',
                        'complement': str(row.get('complement')) if pd.notna(row.get('complement')) else '',
                        'zipCode': zip_code_to_update,
                        'city': str(row.get('city')) if pd.notna(row.get('city')) else '',
                        'country': str(row.get('Country')) if pd.notna(row.get('Country')) else '',
                        "type": "VIS",
                        'source' : source_variable
                    })
                    with open(output_json_file, 'w') as json_file:
                        json.dump(company_json, json_file, indent=4)


            # Else insert address field
            else:
                company_json['address'] = [{'isoCode' : 'en',
                                            "street" : str(row.get('street')) if pd.notna(row.get('street')) else '',
                                            'complement' : str(row.get('complement')) if pd.notna(row.get('complement')) else '',
                                            'zipCode' : str(int(row.get('zipCode'))) if pd.notna(row.get('zipCode')) else '',
                                            'city' : str(row.get('city')) if pd.notna(row.get('city')) else '',
                                            'country' : str(row.get('Country')) if pd.notna(row.get('Country')) else '',
                                            "type": "VIS",
                                            'source' : source_variable
                                            }]
                with open(output_json_file, 'w') as json_file:
                        json.dump(company_json, json_file, indent=4)

    except Exception as e:
        print(f'Exception: {e}')

    # Adding logic for registration number field
    try:
        current_reg = str(row.get('registrations')).strip()
        if pd.notna(row['registrations']):
            
            if "registrations" in company_json:
                registration_numbers = [registration.get("number", "") for registration in company_json.get("registrations", [])]

                if registration_numbers:
                    # Update the existing registration if present
                    company_json['registrations'][0]['number'] = current_reg
                    company_json['registrations'][0]['type'] = "REG"
                    company_json['registrations'][0]['source'] = source_variable
                else:
                    # Add a new registration
                    company_json['registrations'] = [{
                        "type": "REG",
                        "number": current_reg,
                        'source' : source_variable
                    }]
            else:
                # If "registrations" field is not present, create it
                company_json['registrations'] = [{
                    "type": "REG",
                    "number": current_reg,
                    'source' : source_variable
                }]
            
            # Write back to the JSON file
            with open(output_json_file, 'w') as json_file:
                json.dump(company_json, json_file, indent=4)

    except Exception as e:
        print(f'Exception: {e}')

    # Adding logic for VAT codes
    try:
        new_vat_code = str(row.get('vatCode'))
        if pd.notna(row['vatCode']):
    
            if 'vatCode' in company_json:
                # existing_vat_code = company_json.get("vatCode", "") 
                # new_vat_code = row.get('vatCode')
                # print('Existing vat code :',existing_vat_code)
                # print('New vat code :',new_vat_code)
                
                company_json['vatCode'] = new_vat_code
                with open(output_json_file, 'w') as json_file:
                    json.dump(company_json, json_file, indent=4)

            else:
                company_json['vatCode'] = new_vat_code
                
                with open(output_json_file, 'w') as json_file:
                    json.dump(company_json, json_file, indent=4)

    except Exception as e:
        print(f'Exception: {e}')
    
    # Logic for activities
    try:
        if pd.notna(row['activities']):

            if 'activities' in company_json:
                isocodes =  [activity["isoCode"] for activity in company_json.get("activities", [])]
                current_iso_code = "en"
                index_to_update = next((index for index, activity in enumerate(company_json.get("activities", [])) if activity["isoCode"] == "en"), None)
                if index_to_update is not None:
                    company_json["activities"][index_to_update]["value"] = str(row.get('activities'))
                    company_json["activities"][index_to_update]["source"] = source_variable

                if not company_json['activities']:
                    company_json['activities'] = [{"isoCode": "en",
                    'source' : source_variable,
                    "value" :str(row.get('activities'))}]

                with open(output_json_file, 'w') as json_file:
                    json.dump(company_json, json_file, indent=4)


            else:
                company_json['activities'] = [{"isoCode": "en",
                                               'source' : source_variable,
                    "value" :str(row.get('activities'))}]
                
                with open(output_json_file, 'w') as json_file:
                    json.dump(company_json, json_file, indent=4)

    except Exception as e:
        print(f'Exception: {e}')

    # Logic for the exporter fields
    try:
        if pd.notna(row['exporter']):
            current_value = row['exporter']
            if 'exporter' in company_json:
                #current_value = row['exporter']
                #print('Current Exporter value:', current_value)

                if isinstance(current_value, (str, bool)):
                    # If the current value is a string or boolean, convert it to a Python boolean
                    python_boolean = str(current_value).strip().lower() == "true"
                    #print("Converted boolean value is :",python_boolean)
                    company_json['exporter'] = python_boolean
                    with open(output_json_file, 'w') as json_file:
                        json.dump(company_json, json_file, indent=4)
                elif isinstance(current_value, (int, float)):
                    # If the current value is a numeric type, convert it to a Python boolean
                    python_boolean = bool(current_value)
                    #print("Converted boolean value is :",python_boolean)
                    company_json['exporter'] = python_boolean
                    with open(output_json_file, 'w') as json_file:
                        json.dump(company_json, json_file, indent=4)
            else:
                # If 'exporter' is not in the JSON, create it
                # current_value = row['exporter']
                #print('Current Exporter value:', current_value)

                if isinstance(current_value, (str, bool)):
                    # If the current value is a string or boolean, convert it to a Python boolean
                    python_boolean = str(current_value).strip().lower() == "true"
                    #print("Converted boolean value is :",python_boolean)
                    company_json['exporter'] = python_boolean
                    with open(output_json_file, 'w') as json_file:
                        json.dump(company_json, json_file, indent=4)
                elif isinstance(current_value, (int, float)):
                    # If the current value is a numeric type, convert it to a Python boolean
                    python_boolean = bool(current_value)
                    #print("Converted boolean value is :",python_boolean)
                    company_json['exporter'] = python_boolean
                    with open(output_json_file, 'w') as json_file:
                        json.dump(company_json, json_file, indent=4)
                
                # company_json['exporter'] = python_boolean

                # with open(output_json_file, 'w') as json_file:
                #     json.dump(company_json, json_file, indent=4)

    except Exception as e:
        print(f'Exception: {e}')

    # Logic for importer field
    try:
        if pd.notna(row['importer']):
            current_importer_value = row['importer']
            if 'importer' in company_json:
                # current_importer_value = row['importer']
                #print('Current importer value:', current_importer_value)

                if isinstance(current_importer_value, (str, bool)):
                    # If the current value is a string or boolean, convert it to a Python boolean
                    importer_boolean = str(current_importer_value).strip().lower() == "true"
                    #print("Converted boolean value is :",importer_boolean)
                    company_json['importer'] = importer_boolean
                    
                    with open(output_json_file, 'w') as json_file:
                        json.dump(company_json, json_file, indent=4)

                elif isinstance(current_importer_value, (int, float)):
                    # If the current value is a numeric type, convert it to a Python boolean
                    importer_boolean = bool(current_importer_value)
                    #print("Converted boolean value is :",importer_boolean)
                    company_json['importer'] = importer_boolean
                    
                    with open(output_json_file, 'w') as json_file:
                        json.dump(company_json, json_file, indent=4)
                        
            else:
                # If 'importer' is not in the JSON, create it
                # current_importer_value = row['importer']
                #print('Current importer value:', current_importer_value)

                if isinstance(current_importer_value, (str, bool)):
                    # If the current value is a string or boolean, convert it to a Python boolean
                    importer_boolean = str(current_importer_value).strip().lower() == "true"
                    #print("Converted boolean value is :",importer_boolean)
                    company_json['importer'] = importer_boolean
                    
                    with open(output_json_file, 'w') as json_file:
                        json.dump(company_json, json_file, indent=4)

                elif isinstance(current_importer_value, (int, float)):
                    # If the current value is a numeric type, convert it to a Python boolean
                    importer_boolean = bool(current_importer_value)
                    #print("Converted boolean value is :",importer_boolean)
                    company_json['importer'] = importer_boolean
                    
                    with open(output_json_file, 'w') as json_file:
                        json.dump(company_json, json_file, indent=4)
               
                # company_json['importer'] = importer_boolean

                # with open(output_json_file, 'w') as json_file:
                #     json.dump(company_json, json_file, indent=4)

    except Exception as e:
        print(f'Exception: {e}')

    # Logic for producer fields   
    try:
        if pd.notna(row["producer"]):
            current_producer_value = row["producer"]
            if "producer" in company_json:
                #print('Current producer value:', current_producer_value)

                if isinstance(current_producer_value, (str, bool)):
                    # If the current value is a string or boolean, convert it to a Python boolean
                    producer_boolean = str(current_producer_value).strip().lower() == "true"
                    #print("Converted boolean value is:", producer_boolean)
                    company_json["producer"] = producer_boolean

                    with open(output_json_file, 'w') as json_file:
                        json.dump(company_json, json_file, indent=4)

                elif isinstance(current_producer_value, (int, float)):
                    # If the current value is a numeric type, convert it to a Python boolean
                    producer_boolean = bool(current_producer_value)
                    #print('Entered elif producer condition')
                    #print("Converted boolean value is:", producer_boolean)
                    company_json["producer"] = producer_boolean

                    with open(output_json_file, 'w') as json_file:
                        json.dump(company_json, json_file, indent=4)

                    


            else:
                # If 'producer' is not in the JSON, create it
                #print('Current producer value:', current_producer_value)

                if isinstance(current_producer_value, (str, bool)):
                    # If the current value is a string or boolean, convert it to a Python boolean
                    producer_boolean = str(current_producer_value).strip().lower() == "true"
                    #print("Converted boolean value is:", producer_boolean)
                    company_json['producer'] = producer_boolean


                elif isinstance(current_producer_value, (int, float)):
                    # If the current value is a numeric type, convert it to a Python boolean
                    producer_boolean = bool(current_producer_value)
                    #print("Converted boolean value is:", producer_boolean)
                    company_json['producer'] = producer_boolean


                with open(output_json_file, 'w') as json_file:
                            json.dump(company_json, json_file, indent=4)
    

    except Exception as e:
        print(f'Exception: {e}')

    # Logic for distributer field
    try:
        if pd.notna(row['distributor']):
            current_distributor_value = row['distributor']
            if 'distributor' in company_json:
                #print('Current distributor value:', current_distributor_value)

                if isinstance(current_distributor_value, (str, bool)):
                    # If the current value is a string or boolean, convert it to a Python boolean
                    distributor_boolean = str(current_distributor_value).strip().lower() == "true"
                    #print("Converted boolean value is:", distributor_boolean)
                    company_json['distributor'] = distributor_boolean

                    with open(output_json_file, 'w') as json_file:
                        json.dump(company_json, json_file, indent=4)

                elif isinstance(current_distributor_value, (int, float)):
                    # If the current value is a numeric type, convert it to a Python boolean
                    distributor_boolean = bool(current_distributor_value)
                    #print("Converted boolean value is:", distributor_boolean)
                    company_json['distributor'] = distributor_boolean

                    with open(output_json_file, 'w') as json_file:
                        json.dump(company_json, json_file, indent=4)


            else:
                # If 'distributor' is not in the JSON, create it
                #print('Current distributor value:', current_distributor_value)

                if isinstance(current_distributor_value, (str, bool)):
                    # If the current value is a string or boolean, convert it to a Python boolean
                    distributor_boolean = str(current_distributor_value).strip().lower() == "true"
                    #print("Converted boolean value is:", distributor_boolean)
                    company_json['distributor'] = distributor_boolean

                    with open(output_json_file, 'w') as json_file:
                        json.dump(company_json, json_file, indent=4)

                elif isinstance(current_distributor_value, (int, float)):
                    # If the current value is a numeric type, convert it to a Python boolean
                    distributor_boolean = bool(current_distributor_value)
                    #print("Converted boolean value is:", distributor_boolean)
                    company_json['distributor'] = distributor_boolean

                    with open(output_json_file, 'w') as json_file:
                        json.dump(company_json, json_file, indent=4)

                # company_json['distributor'] = distributor_boolean

                # with open(output_json_file, 'w') as json_file:
                #     json.dump(company_json, json_file, indent=4)
                        
    except Exception as e:
        print(f'Exception: {e}')

    # Logic for Service Provider field
    try:
        if pd.notna(row['serviceProvider']):
            current_service_provider_value = row['serviceProvider']
            if 'serviceProvider' in company_json:
                #print('Current service provider value:', current_service_provider_value)

                if isinstance(current_service_provider_value, (str, bool)):
                    # If the current value is a string or boolean, convert it to a Python boolean
                    service_provider_boolean = str(current_service_provider_value).strip().lower() == "true"
                    #print("Converted boolean value is:", service_provider_boolean)
                    company_json['serviceProvider'] = service_provider_boolean

                    with open(output_json_file, 'w') as json_file:
                        json.dump(company_json, json_file, indent=4)

                elif isinstance(current_service_provider_value, (int, float)):
                    # If the current value is a numeric type, convert it to a Python boolean
                    service_provider_boolean = bool(current_service_provider_value)
                    #print("Converted boolean value is:", service_provider_boolean)
                    company_json['serviceProvider'] = service_provider_boolean

                    with open(output_json_file, 'w') as json_file:
                        json.dump(company_json, json_file, indent=4)

            else:
                # If 'serviceProvider' is not in the JSON, create it
                #print('Current service provider value:', current_service_provider_value)

                if isinstance(current_service_provider_value, (str, bool)):
                    # If the current value is a string or boolean, convert it to a Python boolean
                    service_provider_boolean = str(current_service_provider_value).strip().lower() == "true"
                    #print("Converted boolean value is:", service_provider_boolean)
                    company_json['serviceProvider'] = service_provider_boolean

                    with open(output_json_file, 'w') as json_file:
                        json.dump(company_json, json_file, indent=4)

                elif isinstance(current_service_provider_value, (int, float)):
                    # If the current value is a numeric type, convert it to a Python boolean
                    service_provider_boolean = bool(current_service_provider_value)
                    #print("Converted boolean value is:", service_provider_boolean)
                    company_json['serviceProvider'] = service_provider_boolean

                    with open(output_json_file, 'w') as json_file:
                        json.dump(company_json, json_file, indent=4)


    except Exception as e:
        print(f'Exception: {e}')

    
    # Logic for eCommerce field
        
    try:
        if pd.notna(row['eCommerce']):
            current_ecommerce_value = row['eCommerce']
            if 'eCommerce' in company_json:
                #print('Current eCommerce value:', current_ecommerce_value)

                if isinstance(current_ecommerce_value, (str, bool)):
                    # If the current value is a string or boolean, convert it to a Python boolean
                    ecommerce_boolean = str(current_ecommerce_value).strip().lower() == "true"
                    #print("Converted boolean value is:", ecommerce_boolean)
                    company_json['eCommerce'] = ecommerce_boolean

                    with open(output_json_file, 'w') as json_file:
                        json.dump(company_json, json_file, indent=4)

                elif isinstance(current_ecommerce_value, (int, float)):
                    # If the current value is a numeric type, convert it to a Python boolean
                    ecommerce_boolean = bool(current_ecommerce_value)
                    #print("Converted boolean value is:", ecommerce_boolean)
                    company_json['eCommerce'] = ecommerce_boolean

                    with open(output_json_file, 'w') as json_file:
                        json.dump(company_json, json_file, indent=4)
            else:
                ecommerce_boolean = bool(current_ecommerce_value)
                company_json['eCommerce'] = ecommerce_boolean

                with open(output_json_file, 'w') as json_file:
                    json.dump(company_json, json_file, indent=4)

    except Exception as e:
        print(f'Exception: {e}')


    # Logic for LegalForm
    
    try:
        if pd.notna(row['legalForm']):
            if "legalForm" in company_json:
                company_json['legalForm'] = str(row.get('legalForm'))

                with open(output_json_file, 'w') as json_file:
                    json.dump(company_json, json_file, indent=4)

            else:
                company_json['legalForm'] = str(row.get('legalForm'))
                with open(output_json_file, 'w') as json_file:
                    json.dump(company_json, json_file, indent=4)

    except Exception as e:
        print(f'Exception: {e}')

    # Logic for export countries
    try:
        if pd.notna(row['exportCountries']):
            current_countries_list = str(row.get('exportCountries')).strip()
            if 'exportCountries' in company_json:
                existing_export_countries = set(country.lower().strip() for country in company_json.get("exportCountries", []))
                new_export_countries = [country.strip().strip('"') for country in current_countries_list.split(',')]

                # Remove 'nan' and duplicates while ignoring case
                new_export_countries = set(country.lower() for country in new_export_countries if country.lower() != 'nan')

                # Add only new countries to the existing set
                new_countries = new_export_countries - existing_export_countries

                # Convert countries to uppercase before updating the JSON
                new_countries_uppercase = [country.upper() for country in new_countries]
                company_json.setdefault("exportCountries", []).extend(new_countries_uppercase)
                company_json['exportCountries'] = [
                country.upper() for country in company_json.get('exportCountries', []) if country.lower() not in ['nan', 'none']
                ]       
                with open(output_json_file, 'w') as json_file:
                    json.dump(company_json, json_file, indent=4)

                # print('Export Countries:', company_json['exportCountries'])
            else:
                # If 'exportCountries' is not in the JSON, create it with new countries
                new_export_countries = [country.strip().strip('"') for country in current_countries_list.split(',')]
                new_export_countries = set(country.lower() for country in new_export_countries if country.lower() != 'nan')
                new_countries_uppercase = [country.upper() for country in new_export_countries]
                company_json['exportCountries'] = list(new_countries_uppercase)
                company_json['exportCountries'] = [
                 country.upper() for country in company_json.get('exportCountries', []) if country.lower() not in ['nan', 'none']
                ]

                with open(output_json_file, 'w') as json_file:
                    json.dump(company_json, json_file, indent=4)

    except Exception as e:
        print(f'Exception: {e}')

    # Logic for import countries
    
    try:
        if pd.notna(row['importCountries']):
            current_countries_list = str(row.get('importCountries')).strip()
            if 'importCountries' in company_json:
                existing_import_countries = set(country.lower().strip() for country in company_json.get("importCountries", []))
                new_import_countries = [country.strip().strip('"') for country in current_countries_list.split(',')]

                # Remove 'nan' and duplicates while ignoring case
                new_import_countries = set(country.lower() for country in new_import_countries if country.lower() != 'nan')

                # Add only new countries to the existing set
                new_countries = new_import_countries - existing_import_countries

                # Convert countries to uppercase before updating the JSON
                new_countries_uppercase = [country.upper() for country in new_countries]
                company_json.setdefault("importCountries", []).extend(new_countries_uppercase)
                company_json['importCountries'] = [
                country.upper() for country in company_json.get('importCountries', []) if country.lower() not in ['nan', 'none']
                ]
                with open(output_json_file, 'w') as json_file:
                    json.dump(company_json, json_file, indent=4)

                # print('Import Countries:', company_json['importCountries'])
            else:
                # If 'importCountries' is not in the JSON, create it with new countries
                new_import_countries = [country.strip().strip('"') for country in current_countries_list.split(',')]
                new_import_countries = set(country.lower() for country in new_import_countries if country.lower() != 'nan')
                new_countries_uppercase = [country.upper() for country in new_import_countries]
                company_json['importCountries'] = list(new_countries_uppercase)
                company_json['importCountries'] = [
                country.upper() for country in company_json.get('importCountries', []) if country.lower() not in ['nan', 'none']
                ]
                with open(output_json_file, 'w') as json_file:
                    json.dump(company_json, json_file, indent=4)

    except Exception as e:
        print(f'Exception: {e}')

    # Logic for export areas
    try:
        if pd.notna(row['exportAreas']):
            current_areas_list = str(row.get('exportAreas')).strip()
            new_export_areas = {area.strip().upper() for area in current_areas_list.split(',')}
            new_export_areas.discard('NAN')

            # Remove duplicates and strip whitespace
            new_export_areas = {type.strip() for type in new_export_areas}

            if 'exportAreas' in company_json:
                existing_export_areas = {area.strip().upper() for area in company_json.get("exportAreas", [])}
                new_areas = new_export_areas - existing_export_areas
                company_json.setdefault("exportAreas", []).extend(new_areas)
                company_json['exportAreas'] = list(set(company_json['exportAreas']))
                with open(output_json_file, 'w') as json_file:
                    json.dump(company_json, json_file, indent=4)
            else:
                company_json['exportAreas'] = list(new_export_areas)

                # Strip whitespace from the final list
                company_json['exportAreas'] = [type.strip() for type in company_json['exportAreas']]

                company_json['exportAreas'] = list(set(company_json['exportAreas']))

            with open(output_json_file, 'w') as json_file:
                json.dump(company_json, json_file, indent=4)

        

    except Exception as e:
        print(f'Exception: {e}')

    
    # Logic for import areas
    try:
        if pd.notna(row['importAreas']):
            current_areas_list = str(row.get('importAreas')).strip()
            new_import_areas = {area.strip().upper() for area in current_areas_list.split(',')}
            new_import_areas.discard('NAN')

            # Remove duplicates and strip whitespace
            new_import_areas = {area.strip() for area in new_import_areas}

            if 'importAreas' in company_json:
                existing_import_areas = {area.strip().upper() for area in company_json.get("importAreas", [])}
                new_areas = new_import_areas - existing_import_areas
                company_json.setdefault("importAreas", []).extend(new_areas)
            else:
                company_json['importAreas'] = list(new_import_areas)

            # Strip whitespace from the final list
            company_json['importAreas'] = [area.strip() for area in company_json['importAreas']]

            with open(output_json_file, 'w') as json_file:
                json.dump(company_json, json_file, indent=4)

    except Exception as e:
        print(f'Exception: {e}')


    # Logic for INKC field
   
    try:

        
        if pd.notna(row['inkc']):
            inkc_value = int(row.get('inkc'))
            if 'inkc' in company_json:

                if(inkc_value==1 or inkc_value==0 or inkc_value==2):
                    company_json['inkc'] = inkc_value
                    with open(output_json_file, 'w') as json_file:
                        json.dump(company_json, json_file, indent=4)

            else:
                if(inkc_value==1 or inkc_value==0 or inkc_value==2):
                    company_json['inkc'] = inkc_value
                    with open(output_json_file, 'w') as json_file:
                        json.dump(company_json, json_file, indent=4)



    except Exception as e:
        print(f'Exception: {e}')

    # Logic for iebol field
        
    try:
        if pd.notna(row['iebol']):
            iebol_value = int(row.get('iebol'))
            if 'iebol' in company_json:

                if iebol_value == 1 or iebol_value == 0 or iebol_value == 2:
                    company_json['iebol'] = iebol_value
                    with open(output_json_file, 'w') as json_file:
                        json.dump(company_json, json_file, indent=4)
            
            else:
                if iebol_value == 1 or iebol_value == 0 or iebol_value == 2:
                    company_json['iebol'] = iebol_value
                    with open(output_json_file, 'w') as json_file:
                        json.dump(company_json, json_file, indent=4)

    except Exception as e:
        print(f'Exception: {e}')

    # Logic for establishment types
    
    try:
        if pd.notna(row['establishmentTypes']):
            company_json['establishmentTypes'] = [type.strip() for type in company_json.get('establishmentTypes', [])] if company_json.get('establishmentTypes') else []  # Strip all values
            current_types_list = str(row.get('establishmentTypes')).strip()
            if 'establishmentTypes' in company_json:
                existing_types = {type.strip().upper() for type in company_json.get("establishmentTypes", [])}

                # Remove whitespace from existing types
                existing_types = {type.strip() for type in existing_types}

                new_types = {type.strip().strip('"').upper() for type in current_types_list.split(',')}

                # Remove 'nan' and duplicates
                new_types.discard('NAN')

                # Add only new types to the existing set
                new_types_set = new_types - existing_types

                company_json.setdefault("establishmentTypes", []).extend(new_types_set)
                company_json['establishmentTypes'] = list(company_json['establishmentTypes'])

                with open(output_json_file, 'w') as json_file:
                    json.dump(company_json, json_file, indent=4)

                # print('Establishment Types:', company_json['establishmentTypes'])
            else:
                # If 'establishmentTypes' is not in the JSON, create it with new types
                new_types = {type.strip().strip('"').upper() for type in current_types_list.split(',')}
                new_types.discard('NAN')

                company_json['establishmentTypes'] = list(new_types)

                company_json['establishmentTypes'] = [type.strip() for type in company_json['establishmentTypes']]  # Strip all values

                with open(output_json_file, 'w') as json_file:
                    json.dump(company_json, json_file, indent=4)

    except Exception as e:
        print(f'Exception: {e}')


    # Logic for creation dates
    
    try:
        if pd.notna(row['creationDate']):
            if "creationDate" in company_json:
                current_year = int(row.get('creationDate'))
                current_year = datetime.datetime(current_year,1,1)
                current_year = current_year.strftime('%Y-%m-%dT%H:%M:%S.000Z')
                #print('Current year :',current_year)

                company_json['creationDate'] = current_year

                with open(output_json_file, 'w') as json_file:
                    json.dump(company_json, json_file, indent=4)
            
            else:
                current_year = int(row.get('creationDate'))
                current_year = datetime.datetime(current_year,1,1)
                current_year = current_year.strftime('%Y-%m-%dT%H:%M:%S.000Z')
                print('Current year :',current_year)
                company_json['creationDate'] = current_year

                with open(output_json_file, 'w') as json_file:
                    json.dump(company_json, json_file, indent=4)

    except Exception as e:
        print(f'Exception: {e}')

    # Logic for Banks field 
    try:
        if pd.notna(row['banks']):
            current_bank = custom_title_case(str(row.get('banks')).strip())

            if 'banks' in company_json:
                existing_banks = {bank["name"].strip().lower() for bank in company_json.get("banks", [])}
                #print('Existing banks are:', existing_banks)

                if current_bank.lower() not in existing_banks:
                    company_json['banks'].append({
                        'isoCode': 'en',
                        'name': current_bank
                    })

                with open(output_json_file, 'w') as json_file:
                    json.dump(company_json, json_file, indent=4)
            else:
                company_json['banks'] = [{
                    'isoCode': 'en',
                    'name': current_bank
                }]
                with open(output_json_file, 'w') as json_file:
                    json.dump(company_json, json_file, indent=4)

    except Exception as e:
        print(f'Exception: {e}')

    # Logic for associations field
    try:
        if pd.notna(row['associations']):
            current_association = custom_title_case(str(row.get('associations')).strip())

            if 'associations' in company_json:
                existing_associations = {association["name"].strip().lower() for association in company_json.get("associations", [])}
                # print('Existing associations are:', existing_associations)

                if current_association.lower() not in existing_associations:
                    company_json['associations'].append({
                        'isoCode': 'en',
                        'name': current_association
                    })

                with open(output_json_file, 'w') as json_file:
                    json.dump(company_json, json_file, indent=4)
            else:
                company_json['associations'] = [{
                    'isoCode': 'en',
                    'name': current_association
                }]
                with open(output_json_file, 'w') as json_file:
                    json.dump(company_json, json_file, indent=4)

    except Exception as e:
        print(f'Exception: {e}')

    # Logic for othersLocations
        
    try:
        if pd.notna(row['othersLocations']):
            current_location = custom_title_case(str(row.get('othersLocations')).split(':')[0].strip())

            if 'othersLocations' in company_json:
                existing_locations = {location["description"].split(':')[0].strip().lower() for location in company_json.get("othersLocations", [])}
                #print('Existing locations are:', existing_locations)

                if current_location.lower() not in existing_locations:
                    company_json['othersLocations'].append({
                        'description': current_location,
                        'isoCode':'en'
                    })

                with open(output_json_file, 'w') as json_file:
                    json.dump(company_json, json_file, indent=4)
            else:
                company_json['othersLocations'] = [{
                    'description': current_location,
                    'isoCode':'en'
                }]
                with open(output_json_file, 'w') as json_file:
                    json.dump(company_json, json_file, indent=4)

    except Exception as e:
        print(f'Exception: {e}')
    
    # Logic for Certifications
    # try:
    #     if pd.notna(row['certifications']):
    #         current_certification = custom_title_case(str(row.get('certifications')).strip())

    #         if 'certifications' in company_json:
    #             existing_certifications = {certification["name"].strip().lower() for certification in company_json.get("certifications", [])}
    #             # print('Existing certifications are:', existing_certifications)

    #             if current_certification.lower() not in existing_certifications:
    #                 company_json['certifications'].append({
    #                     'isoCode': 'en',
    #                     'name': current_certification,
    #                     'number' : str(row.get('certification_number')),
    #                     'certification_description' : str(row.get('certification_description'))
    #                 })

    #             with open(output_json_file, 'w') as json_file:
    #                 json.dump(company_json, json_file, indent=4)
    #         else:
    #             company_json['certifications'] = [{
    #                 'isoCode': 'en',
    #                 'name': current_certification,
    #                 'number' : str(row.get('certification_number')),
    #                 'certification_description' : str(row.get('certification_description'))
    #             }]
    #             with open(output_json_file, 'w') as json_file:
    #                 json.dump(company_json, json_file, indent=4)

    # except Exception as e:
    #     print(f'Exception: {e}')

    # Logic for Shared Capital
    try:
        if pd.notna(row['SharedCapital']):
            # If 'SharedCapital' is present in the JSON, update its value
            if 'SharedCapital' in company_json:
                company_json['SharedCapital']['value'] = int(row['SharedCapital'])
                company_json['SharedCapital']['source'] = source_variable

                with open(output_json_file, 'w') as json_file:
                    json.dump(company_json, json_file, indent=4)
            else:
                # If 'SharedCapital' is not present, add it
                company_json['SharedCapital'] = {
                    'value': int(row['SharedCapital']),
                    'currency': 'INR',  # Assuming the currency is always INR,
                    'source' : source_variable
                }

            # Write back to the JSON file
            with open(output_json_file, 'w') as json_file:
                json.dump(company_json, json_file, indent=4)

    except Exception as e:
        print(f'Exception: {e}')
    

    # Logic for FaxCodes
    
    # try:
    #     if pd.notna(row['fax_number']):
    #         print('Entered Fax No logic')
    #         if 'fax' in company_json:
    #             existing_fax = company_json.get('fax', {})
    #             existing_fax_no = unicodedata.normalize('NFC', ''.join(char for char in str(existing_fax.get('number')) if char.isdigit()))
    #             current_fax_no = unicodedata.normalize('NFC', str(int(row.get('fax_number'))))
    #             print('Existing Fax No:', existing_fax_no)
    #             print('Current Fax No:', current_fax_no)

    # except Exception as e:
    #     print(f'Exception: {e}')
        
    # Logic for Classifications
    try:
        if pd.notna(row['classifications']) or pd.notna(row['company_type']):
            existing_classification_codes = [classification["code"].strip() for classification in company_json.get("classifications", [])]
            new_classification_codes = [code.strip() for code in str(row['classifications']).split(',') if code.strip()]
            company_type = [ctype.strip() for ctype in str(row['company_type']).split(',') if ctype.strip()] if not pd.isna(row['company_type']) else []
            accepted_company_types = ['producer', 'distributor', 'service', 'export', 'import', '']

            if 'classifications' in company_json:
                for code in new_classification_codes:
                    if code not in existing_classification_codes:
                        # print('Code not in existing classifications code !')
                        valid_company_types = [ctype for ctype in company_type if ctype in accepted_company_types]
                        if valid_company_types:
                            new_classification = {"code": code,'source':source_variable}
                            for type_elem in valid_company_types:
                                new_classification[type_elem] = True
                            company_json["classifications"].append(new_classification)

                    else:
                        # Update existing code with new company types
                        existing_classification = next((c for c in company_json["classifications"] if c["code"] == code), None)
                        if existing_classification:
                            valid_company_types = [ctype for ctype in company_type if ctype in accepted_company_types]
                            for type_elem in valid_company_types:
                                existing_classification[type_elem] = True
                                
            else:
                # If 'classifications' doesn't exist in company_json, create it and append new codes
                company_json["classifications"] = []
                for code in new_classification_codes:
                    # Add the classification if there are any new classification codes
                    new_classification = {"code": code, "source": source_variable}
                    # Add company types if they exist
                    for type_elem in company_type:
                        new_classification[type_elem] = True
                    company_json["classifications"].append(new_classification)

                # Filter out duplicates based on the "code" field
                seen_codes = set()
                company_json["classifications"] = [c for c in company_json["classifications"] if not (c["code"] in seen_codes or seen_codes.add(c["code"]))]


            with open(output_json_file, 'w') as json_file:
                json.dump(company_json, json_file, indent=4)

    except Exception as e:
        print(f'Exception: {e}')




def main():
    excel_file_path = r"D:\VISHAL KOMPASS\web-scraping\Linkedin Company Prompt and Scrape Executives data\Salutation_LinkedinExecutives_PerCompanyPrompt_19-03-2025.xlsx"
    backup_directory = r"C:\Users\Kompass\Desktop"
    output_directory = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
    os.makedirs(output_directory, exist_ok=True)  # Create the output directory if it doesn't exist
    output_json_file = os.path.join(output_directory, "output.json")
    session_cookie = r"connect.sid=s%3Axinpmg7ACxfaYPqBVTxU-ZlO2JiGp7QH.wkaZ8GgvOTVBlafB9A8ITFFfoOEogQ8rXEWEqr9JtC0"

    print('Output Json path :',output_json_file)
    print()
    excel_file = read_excel(excel_file_path)
    print('Excel File Start :')

    print(excel_file.head())
    print()
    print('Excel File End :')
    print(excel_file.tail())
    print()
    print("Columns present in the excel :")
    print()

    for i, column_name in enumerate(excel_file.columns, start=1):
        print(f"{i} - {column_name}")

    print()
    print('Columns not present in the excel :')
    print()
    
    # # Add a new column for Response Status
    # excel_file['Response Status'] = None

    total_rows = len(excel_file)
    for index, row in tqdm(excel_file.iterrows(),desc='Progress Bar : ',total=total_rows):
        company_id = row['_id']

        try:
            response = get_comp_json(company_id, session_cookie)
            if response and response.status_code != 200:
                raise Exception(f"Failed to get JSON. Status Code: {response.status_code}")
            json_data = response.json()
        except Exception as e:
            print(f"Failed to fetch data for company {company_id}. Error: {e}")
            # To create a new company
            # json_data = {"companyId": str(company_id).strip()}  # Initialize with an empty dictionary if the get request fails.
            # print("Initialized JSON is:", json_data)

        # backup_file_path = os.path.join(backup_directory, f"{company_id}.json")
        # write_to_json(json_data, backup_file_path)
        print()
        update_comp_json(row, json_data, output_json_file)
        # response_status = update_comp_json(row, json_data, output_json_file)
        # excel_file.at[index, 'Response Status'] = response_status  # Update response status in the DataFrame

        try:
            post_comp_json(company_id, output_json_file, session_cookie)
        except Exception as e:
            print(f"Failed to post data for company {company_id}. Error: {e}")
            
     # Save the updated Excel file with Response Status
    # updated_excel_path = os.path.join(r"C:\Users\Kompass\Desktop\Hunter IO Files\Status_Test.xlsx")
    # excel_file.to_excel(updated_excel_path, index=False)
    # print(f"Updated Excel file saved at: {updated_excel_path}")

if __name__ == "__main__":
    main()