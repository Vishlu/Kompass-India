
# from firecrawl import FirecrawlApp

# # Initialize the FirecrawlApp with your API key

# app = FirecrawlApp(api_key='fc-65765ac1e3de4325b4728fe6881c6409')

# data = app.extract([
# '20microns.com'
# ], {
# 'prompt': "Extract the following details from the webpage: Company Name, website links, City, Country, Address in street and complements, Email Id, phone no., description, product and services",
# 'enableWebSearch': True # Enable web search for better context
# })
# print(data)




import pandas as pd
from firecrawl import FirecrawlApp
from tqdm import tqdm  # Import tqdm for the progress bar

# Initialize the FirecrawlApp with your API key
app = FirecrawlApp(api_key='fc-65765ac1e3de4325b4728fe6881c6409')

# Read the input Excel file containing website URLs and the _id column
input_file = r"C:\Users\Kompass\Downloads\Full Data.xlsx"  # Replace with your input file path
df_input = pd.read_excel(input_file, usecols=['webSites', '_id'])  # Include the '_id' column

# Prepare a list to store the extracted data
extracted_data = []

# Initialize the progress bar
total_rows = len(df_input)
with tqdm(total=total_rows, desc="Processing websites", unit="website") as pbar:
    # Iterate over each website URL in the Excel file
    for index, row in df_input.iterrows():
        urls = row['webSites'].split(',')  # Assuming multiple URLs are comma-separated
        for url in urls:
            url = url.strip()  # Remove any leading/trailing whitespace
            try:
                # Extract data using the Firecrawl API
                data = app.extract([url], {
                    'prompt': "Extract the following details from the webpage: "
                              "Company Name, website links, City, Country, Address in street and complements, "
                              "Email Id, phone no., description, product and services",
                    'enableWebSearch': True  # Enable web search for better context
                })

                # Check if the extraction was successful
                if data.get('success', False):
                    extracted_info = data['data']

                    # Add the website URL, extracted info, and _id to the data list
                    extracted_data.append({
                        '_id': row['_id'],  # Add _id to the output
                        'webSite': url,
                        'Company Name': extracted_info.get('companyName', ''),
                        'Website Links': ', '.join(extracted_info.get('websiteLinks', [])),
                        'City': extracted_info.get('city', ''),
                        'Country': extracted_info.get('country', ''),
                        'Street': extracted_info.get('address', {}).get('street', ''),
                        'Complement': extracted_info.get('address', {}).get('complements', ''),
                        'Email ID': extracted_info.get('emailId', ''),
                        'Phone Number': extracted_info.get('phoneNo', ''),
                        'Description': extracted_info.get('description', ''),
                        'Products and Services': extracted_info.get('productsAndServices', ''),
                    })
                else:
                    print(f"Extraction failed for {url}: {data.get('error', 'Unknown error')}")

            except Exception as e:
                print(f"An error occurred for {url}: {e}")

            # Update the progress bar
            pbar.update(1)

# Convert the extracted data into a DataFrame
df_output = pd.DataFrame(extracted_data)

# Write the output DataFrame to an Excel file
output_file = r"C:\Users\Kompass\Desktop\Email Generator 22-01-2025\FireCrawl_Marketing_Manager_EmailVerification_InputFile_23.01.2025.xlsx"  # Replace with your desired output file path
df_output.to_excel(output_file, index=False)

print(f"Data extraction complete. Output saved to {output_file}.")
