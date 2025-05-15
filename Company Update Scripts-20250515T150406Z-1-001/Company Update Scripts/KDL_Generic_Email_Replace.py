"""
This script is used to clean all exiting emails and replace with new emails on the basis on _id which we give through the excel sheet.

Below is the Example of input file:  

_id	        emails
INZ9999999	hq@aakashpackaging.com
INZ9999999  aacharya@mahalaxmidhatu.com,qnd@mahalaxmidhatu.com,sales@mahalaxmidhatu.com
"""


# # Importing the necessary libraries
# import pandas as pd
# import json
# import requests
# from pprint import pprint
# import os
# import datetime
# import unicodedata
# from tqdm import tqdm

# # Get current date
# current_date = datetime.datetime.now()
# print('Current date is :', current_date)

# # Format the date components as required (KDL_day_month_year)
# formatted_date = current_date.strftime("KDL_%d_%m_%Y")

# # Assign the formatted date to your source variable
# source_variable = str(formatted_date).strip()

# # Function to read Excel and strip whitespace
# def read_excel(file_path):
#     df = pd.read_excel(file_path)
#     df = df.apply(lambda x: x.strip() if isinstance(x, str) else x)
#     df.columns = [col.strip() if isinstance(col, str) else col for col in df.columns]
#     return df

# # Function to get Company JSON from database
# def get_comp_json(company_id, session_cookie):
#     base_url = "https://translations.kompass-int.com/kompass/companies/"
#     url = base_url + company_id.strip('"')

#     headers = {
#         'User-Agent': 'Mozilla/5.0',
#         'Cookie': session_cookie
#     }

#     try:
#         response = requests.get(url, auth=('kin', 'U&8#JjnUMzR&Yj'), headers=headers)
#         response.raise_for_status()  # Raise error for HTTP failures (4xx, 5xx)
#         return response.json()
#     except requests.exceptions.RequestException as e:
#         print(f'Error fetching data for company {company_id}: {e}')
#         return None

# # Function to write to JSON file
# def write_to_json(json_data, file_path):
#     with open(file_path, 'w', encoding='utf-8') as json_file:
#         json.dump(json_data, json_file, indent=4)

# # Function to post updated data to Kompass DB
# def post_comp_json(company_id, updated_json, session_cookie):
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
#             print(f"Company {company_id} updated successfully in Kompass DB.")
#         else:
#             print(f"Failed to update company {company_id}. Status Code: {response.status_code} - {response.text}")

#     except requests.exceptions.RequestException as e:
#         print(f"Error updating data for company {company_id}: {e}")


# # Function to update emails in JSON
# def update_emails(email_data, company_json):
#     try:
#         new_emails = []
#         for _, row in email_data.iterrows():
#             if pd.notna(row.get('emails')):
#                 # Split the emails by comma and strip whitespace
#                 emails = [email.strip() for email in row.get('emails').split(',')]
#                 for email in emails:
#                     new_email = {
#                         'value': email,
#                         'source': source_variable
#                     }
#                     new_emails.append(new_email)

#         # Replace old emails with the new list
#         if new_emails:
#             company_json['emails'] = new_emails
#     except Exception as e:
#         print(f'Error updating emails: {e}')

# # Main function
# def main():
#     excel_file_path = r"C:\Users\Kompass\Desktop\Indian Website\Email Batch Upload\Stack_Batch(2001-4000)Email_Upload_4.3.2025.xlsx"
#     output_directory = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
#     os.makedirs(output_directory, exist_ok=True)
#     output_json_file = os.path.join(output_directory, "output.json")
#     session_cookie = r"connect.sid=s%3Axinpmg7ACxfaYPqBVTxU-ZlO2JiGp7QH.wkaZ8GgvOTVBlafB9A8ITFFfoOEogQ8rXEWEqr9JtC0"

#     print('Output Json path :', output_json_file)
#     print()

#     excel_file = read_excel(excel_file_path)
#     if excel_file is None:
#         print("Failed to read the Excel file, exiting.")
#         return

#     print('Excel File Start :')
#     print(excel_file.head())
#     print()
#     print('Excel File End :')
#     print(excel_file.tail())
#     print()

#     # Group rows by _id to handle multiple phone numbers and emails
#     grouped_data = excel_file.groupby('_id')

#     # Process each company
#     for company_id, group in tqdm(grouped_data, desc='Progress Bar : '):
#         print(f"Processing company ID: {company_id}")

#         # Fetch existing company JSON
#         json_data = get_comp_json(company_id, session_cookie)
#         if json_data:
#             print(f"Company {company_id} data retrieved successfully.")

#             # Update emails with all rows for this company
#             update_emails(group, json_data)

#             # Write updated data to the output JSON file
#             write_to_json(json_data, output_json_file)

#             # Post updated JSON to Kompass DB
#             post_comp_json(company_id, json_data, session_cookie)
#         else:
#             print(f"Failed to retrieve company data for ID {company_id}, skipping update.")

# if __name__ == "__main__":
#     main()



"""
This script is used to clean all existing emails and replace them with new emails based on _id provided through an Excel sheet.

"""

# Importing the necessary libraries
import pandas as pd
import json
import requests
import os
import datetime
from tqdm import tqdm

# Get current date
current_date = datetime.datetime.now()
print('Current date is :', current_date)

# Format the date components as required (KDL_day_month_year)
formatted_date = current_date.strftime("KDL_%d_%m_%Y")

# Assign the formatted date to your source variable
source_variable = str(formatted_date).strip()

# Function to read Excel and strip whitespace
def read_excel(file_path):
    df = pd.read_excel(file_path)
    df = df.apply(lambda x: x.strip() if isinstance(x, str) else x)
    df.columns = [col.strip() if isinstance(col, str) else col for col in df.columns]
    return df

# Function to get Company JSON from database
def get_comp_json(company_id, session_cookie):
    base_url = "https://translations.kompass-int.com/kompass/companies/"
    url = base_url + company_id.strip('"')

    headers = {
        'User-Agent': 'Mozilla/5.0',
        'Cookie': session_cookie
    }

    try:
        response = requests.get(url, auth=('kin', 'U&8#JjnUMzR&Yj'), headers=headers)
        response.raise_for_status()  # Raise error for HTTP failures (4xx, 5xx)
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f'Error fetching data for company {company_id}: {e}')
        return None

# Function to write to JSON file
def write_to_json(json_data, file_path):
    with open(file_path, 'w', encoding='utf-8') as json_file:
        json.dump(json_data, json_file, indent=4)

# Function to post updated data to Kompass DB
def post_comp_json(company_id, updated_json, session_cookie):
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
            print(f"Company {company_id} updated successfully in Kompass DB.")
        else:
            print(f"Failed to update company {company_id}. Status Code: {response.status_code} - {response.text}")
    except requests.exceptions.RequestException as e:
        print(f"Error updating data for company {company_id}: {e}")

# Function to update emails in JSON
def update_emails(email_data, company_json):
    try:
        new_emails = []
        for _, row in email_data.iterrows():
            if pd.notna(row.get('emails')):
                # Split the emails by comma and strip whitespace
                emails = [email.strip() for email in row.get('emails').split(',')]
                for email in emails:
                    new_email = {
                        'value': email,
                        'source': source_variable
                    }
                    new_emails.append(new_email)

        # Replace old emails with the new list
        if new_emails:
            company_json['emails'] = new_emails
    except Exception as e:
        print(f'Error updating emails: {e}')

# Main function
def main():
    excel_file_path = r"C:\Users\Kompass\Desktop\New Cleaned Indian Websites\New Email Upload\Batch(28001-30000)Emails_Indian_website_13-3-2025.xlsx"
    output_directory = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
    os.makedirs(output_directory, exist_ok=True)
    output_json_file = os.path.join(output_directory, "output.json")
    session_cookie = r"connect.sid=s%3Axinpmg7ACxfaYPqBVTxU-ZlO2JiGp7QH.wkaZ8GgvOTVBlafB9A8ITFFfoOEogQ8rXEWEqr9JtC0"

    print('Output Json path :', output_json_file)
    print()

    excel_file = read_excel(excel_file_path)
    if excel_file is None:
        print("Failed to read the Excel file, exiting.")
        return

    print('Excel File Start :')
    print(excel_file.head())
    print()
    print('Excel File End :')
    print(excel_file.tail())
    print()

    # Group rows by _id to handle multiple phone numbers and emails
    grouped_data = excel_file.groupby('_id')

    # Process each company
    for company_id, group in tqdm(grouped_data, desc='Processing Companies'):
        print(f"\nProcessing company ID: {company_id}")

        # Fetch existing company JSON
        json_data = get_comp_json(company_id, session_cookie)
        if json_data:
            print(f"Company {company_id} data retrieved successfully.")

            # Update emails with all rows for this company
            update_emails(group, json_data)

            # Write updated data to the output JSON file
            write_to_json(json_data, output_json_file)

            # Post updated JSON to Kompass DB
            post_comp_json(company_id, json_data, session_cookie)
        else:
            print(f"Failed to retrieve company data for ID {company_id}, skipping update.")

if __name__ == "__main__":
    main()