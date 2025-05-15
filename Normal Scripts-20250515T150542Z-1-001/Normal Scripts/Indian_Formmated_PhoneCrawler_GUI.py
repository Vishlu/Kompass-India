
# import pandas as pd
# import requests
# from bs4 import BeautifulSoup
# import re
# from fake_useragent import UserAgent
# from urllib.parse import urljoin

# # Improved regex pattern for Indian and international phone numbers
# INDIAN_PHONE_REGEX = r'(?<!\d)(?:\(\+\d{1,3}\)|(?:\+91|0|91))?[ -]?(?:\d{4,5}[ -]?\d{4,5}|\(?\d{2,4}\)?[ -]?\d{3,4}[ -]?\d{3,4}|\d{10})(?!\d)'

# # Function to normalize phone numbers and filter only Indian numbers
# def normalize_phone_number(phone):
#     digits = re.sub(r'\D', '', phone)
    
#     if digits.startswith('0') or (len(digits) == 10 and not digits.startswith('91')):
#         digits = '+91' + digits[-10:]
#     elif len(digits) > 10 and not digits.startswith('+'):
#         digits = '+' + digits  
    
#     if digits.startswith('+91') and len(digits) == 13:
#         return digits
#     return None  

# # Function to extract and filter phone numbers using regex
# def extract_phone_numbers(text):
#     phone_numbers = set(re.findall(INDIAN_PHONE_REGEX, text))
#     normalized_numbers = set()
#     for number in phone_numbers:
#         normalized = normalize_phone_number(number)
#         if normalized:
#             normalized_numbers.add(normalized)
#     return normalized_numbers

# # Function to generate contact-us page URLs
# def generate_contact_urls(base_url):
#     contact_urls = [
#         urljoin(base_url, path)
#         for path in [
#             "contact", "contact-us", "contactus", "contact.html", "contact.php", "contactus.php", "contact.aspx"
#         ]
#     ]
#     return contact_urls

# # Function to crawl a website and extract phone numbers
# def crawl_website(url, user_agent):
#     try:
#         headers = {"User-Agent": user_agent}
#         response = requests.get(url, headers=headers, timeout=15)
#         response.raise_for_status()
#         soup = BeautifulSoup(response.text, "html.parser")
#         text = soup.get_text()
#         phone_numbers = extract_phone_numbers(text)
#         return phone_numbers
#     except Exception as e:
#         print(f"Error crawling {url}: {e}")
#         return set()

# # Function to crawl homepage and contact variations
# def crawl_all_variations(base_url, user_agent):
#     all_numbers = set()
#     urls_to_crawl = [base_url] + generate_contact_urls(base_url)
    
#     for url in urls_to_crawl:
#         print(f"Crawling: {url}")
#         phone_numbers = crawl_website(url, user_agent)
#         all_numbers.update(phone_numbers)
    
#     return all_numbers

# # Function to process websites from Excel file
# def crawl_websites(input_file, output_file):
#     df = pd.read_excel(input_file)
#     if 'Website' not in df.columns:
#         raise ValueError("Column 'Website' not found in the Excel file.")

#     ua = UserAgent()
#     all_results = []

#     for index, row in df.iterrows():
#         website = row['Website']
#         if not website.startswith(('http://', 'https://')):
#             website = 'https://' + website
        
#         print(f"Processing: {website}")
#         user_agent = ua.random
#         phone_numbers = crawl_all_variations(website, user_agent)

#         for phone_number in phone_numbers:
#             all_results.append({
#                 **row.to_dict(),
#                 'Phone Number': phone_number,
#                 'Count': 1
#             })

#     results_df = pd.DataFrame(all_results)
#     results_df.to_excel(output_file, index=False)
#     print(f"Results saved to {output_file}")

# # Main execution
# if __name__ == "__main__":
#     input_file = r"C:\Users\Kompass\Downloads\Test Websites for Phones_25.02.2025.xlsx"
#     output_file = r"C:\Users\Kompass\Downloads\Second_Output_Test Websites for Phones_25.02.2025.xlsx"

#     try:
#         crawl_websites(input_file, output_file)
#     except Exception as e:
#         print(f"An error occurred: {e}")





# """
# stacking, contact variation
# """

# import pandas as pd
# import requests
# from bs4 import BeautifulSoup
# import re
# from fake_useragent import UserAgent
# from urllib.parse import urljoin
# import phonenumbers
# import tkinter as tk
# from tkinter import filedialog, ttk



# # Improved regex pattern for Indian and international phone numbers
# INDIAN_PHONE_REGEX = r'(?<!\d)(?:\(\+\d{1,3}\)|(?:\+91|0|91))?[ -]?(?:\d{4,5}[ -]?\d{4,5}|\(?\d{2,4}\)?[ -]?\d{3,4}[ -]?\d{3,4}|\d{10})(?!\d)'

# # Function to normalize phone numbers and filter only Indian numbers
# def normalize_phone_number(phone):
#     digits = re.sub(r'\D', '', phone)
    
#     if digits.startswith('0') or (len(digits) == 10 and not digits.startswith('91')):
#         digits = '+91' + digits[-10:]
#     elif len(digits) > 10 and not digits.startswith('+'):
#         digits = '+' + digits  
    
#     if digits.startswith('+91') and len(digits) == 13:
#         return digits
#     return None  

# # Function to extract and filter phone numbers using regex
# def extract_phone_numbers(text):
#     phone_numbers = set(re.findall(INDIAN_PHONE_REGEX, text))
#     normalized_numbers = set()
#     for number in phone_numbers:
#         normalized = normalize_phone_number(number)
#         if normalized:
#             normalized_numbers.add(normalized)
#     return normalized_numbers

# # Function to format phone numbers and detect country
# def format_phone_number(phone):
#     try:
#         parsed_number = phonenumbers.parse(phone, None)  # Auto-detect region
#         if phonenumbers.is_valid_number(parsed_number):
#             formatted_number = phonenumbers.format_number(parsed_number, phonenumbers.PhoneNumberFormat.INTERNATIONAL)
#             country = phonenumbers.region_code_for_number(parsed_number)
#             return formatted_number, country
#         else:
#             return phone, "Unknown"
#     except phonenumbers.phonenumberutil.NumberParseException:
#         return phone, "Unknown"

# # Function to generate contact-us page URLs
# def generate_contact_urls(base_url):
#     contact_urls = [
#         urljoin(base_url, path)
#         for path in [
#             "contact", "contact-us", "contactus", "contact.html", "contact.php", "contactus.php", "contact.aspx"
#         ]
#     ]
#     return contact_urls

# # Function to crawl a website and extract phone numbers
# def crawl_website(url, user_agent):
#     try:
#         headers = {"User-Agent": user_agent}
#         response = requests.get(url, headers=headers, timeout=15)
#         response.raise_for_status()
#         soup = BeautifulSoup(response.text, "html.parser")
#         text = soup.get_text()
#         phone_numbers = extract_phone_numbers(text)
#         return phone_numbers
#     except Exception as e:
#         print(f"Error crawling {url}: {e}")
#         return set()

# # Function to crawl homepage and contact variations
# def crawl_all_variations(base_url, user_agent):
#     all_numbers = set()
#     urls_to_crawl = [base_url] + generate_contact_urls(base_url)
    
#     for url in urls_to_crawl:
#         print(f"Crawling: {url}")
#         phone_numbers = crawl_website(url, user_agent)
#         all_numbers.update(phone_numbers)
    
#     return all_numbers

# # Function to process websites from Excel file
# def crawl_websites(input_file, output_file):
#     df = pd.read_excel(input_file)
#     if 'Website' not in df.columns:
#         raise ValueError("Column 'Website' not found in the Excel file.")

#     ua = UserAgent()
#     all_results = []

#     for index, row in df.iterrows():
#         website = row['Website']
#         if not website.startswith(('http://', 'https://')):
#             website = 'https://' + website
        
#         print(f"Processing: {website}")
#         user_agent = ua.random
#         phone_numbers = crawl_all_variations(website, user_agent)

#         for phone_number in phone_numbers:
#             formatted_number, country = format_phone_number(phone_number)
#             all_results.append({
#                 **row.to_dict(),
#                 'Phone Number': phone_number,
#                 'Formatted Phone Number': formatted_number,
#                 'Country': country,
#                 'Count': 1
#             })

#     results_df = pd.DataFrame(all_results)
#     results_df.to_excel(output_file, index=False)
#     print(f"Results saved to {output_file}")

# # Main execution
# if __name__ == "__main__":
#     input_file = r"C:\Users\Kompass\Downloads\Test Websites for Phones_25.02.2025.xlsx"
#     output_file = r"C:\Users\Kompass\Downloads\Third_Output_Test Websites for Phones_25.02.2025.xlsx"

#     try:
#         crawl_websites(input_file, output_file)
#     except Exception as e:
#         print(f"An error occurred: {e}")





import pandas as pd
import requests
from bs4 import BeautifulSoup
import re
from fake_useragent import UserAgent
from urllib.parse import urljoin
import phonenumbers
import tkinter as tk
from tkinter import filedialog, ttk
import threading
import queue

# Improved regex pattern for Indian and international phone numbers
INDIAN_PHONE_REGEX = r'(?<!\d)(?:\(\+\d{1,3}\)|(?:\+91|0|91))?[ -]?(?:\d{4,5}[ -]?\d{4,5}|\(?\d{2,4}\)?[ -]?\d{3,4}[ -]?\d{3,4}|\d{10})(?!\d)'

# Function to normalize phone numbers and filter only Indian numbers
def normalize_phone_number(phone):
    digits = re.sub(r'\D', '', phone)
    if digits.startswith('0') or (len(digits) == 10 and not digits.startswith('91')):
        digits = '+91' + digits[-10:]
    elif len(digits) > 10 and not digits.startswith('+'):
        digits = '+' + digits
    if digits.startswith('+91') and len(digits) == 13:
        return digits
    return None

# Function to extract and filter phone numbers using regex
def extract_phone_numbers(text):
    phone_numbers = set(re.findall(INDIAN_PHONE_REGEX, text))
    normalized_numbers = set()
    for number in phone_numbers:
        normalized = normalize_phone_number(number)
        if normalized:
            normalized_numbers.add(normalized)
    return normalized_numbers

# Function to format phone numbers and detect country
def format_phone_number(phone):
    try:
        parsed_number = phonenumbers.parse(phone, None)  # Auto-detect region
        if phonenumbers.is_valid_number(parsed_number):
            formatted_number = phonenumbers.format_number(parsed_number, phonenumbers.PhoneNumberFormat.INTERNATIONAL)
            country = phonenumbers.region_code_for_number(parsed_number)
            return formatted_number, country
        else:
            return phone, "Unknown"
    except phonenumbers.phonenumberutil.NumberParseException:
        return phone, "Unknown"

# Function to generate contact-us page URLs
def generate_contact_urls(base_url):
    contact_urls = [
        urljoin(base_url, path)
        for path in [
            "contact", "contact-us", "contactus", "contact.html", "contact.php", "contactus.php", "contact.aspx"
        ]
    ]
    return contact_urls

# Function to crawl a website and extract phone numbers
def crawl_website(url, user_agent):
    try:
        headers = {"User-Agent": user_agent}
        response = requests.get(url, headers=headers, timeout=15)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, "html.parser")
        text = soup.get_text()
        phone_numbers = extract_phone_numbers(text)
        return phone_numbers
    except Exception as e:
        print(f"Error crawling {url}: {e}")
        return set()

# Function to crawl homepage and contact variations
def crawl_all_variations(base_url, user_agent, log_queue):
    all_numbers = set()
    urls_to_crawl = [base_url] + generate_contact_urls(base_url)
    for url in urls_to_crawl:
        log_queue.put(f"Crawling: {url}")
        phone_numbers = crawl_website(url, user_agent)
        all_numbers.update(phone_numbers)
    return all_numbers

# Function to process websites from Excel file
def crawl_websites(input_file, output_file, progress_var, log_queue, total_websites):
    try:
        df = pd.read_excel(input_file)
        if 'Website' not in df.columns:
            raise ValueError("Column 'Website' not found in the Excel file.")
        ua = UserAgent()
        all_results = []
        for index, row in df.iterrows():
            website = row['Website']
            if not website.startswith(('http://', 'https://')):
                website = 'https://' + website
            log_queue.put(f"Processing: {website}")
            user_agent = ua.random
            phone_numbers = crawl_all_variations(website, user_agent, log_queue)
            for phone_number in phone_numbers:
                formatted_number, country = format_phone_number(phone_number)
                all_results.append({
                    **row.to_dict(),
                    'Phone Number': phone_number,
                    'Formatted Phone Number': formatted_number,
                    'Country': country,
                    'Count': 1
                })
            progress_var.set((index + 1) / total_websites * 100)
        results_df = pd.DataFrame(all_results)
        results_df.to_excel(output_file, index=False)
        log_queue.put(f"Results saved to {output_file}")
    except Exception as e:
        log_queue.put(f"An error occurred: {e}")

def browse_input_file():
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    input_entry.delete(0, tk.END)
    input_entry.insert(0, filename)

def browse_output_file():
    filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    output_entry.delete(0, tk.END)
    output_entry.insert(0, filename)

def start_crawl():
    input_file = input_entry.get()
    output_file = output_entry.get()
    if not input_file or not output_file:
        log_queue.put("Please select input and output files.")
        return
    try:
        df = pd.read_excel(input_file)
        total_websites = len(df)
    except Exception as e:
        log_queue.put(f"Error reading input file: {e}")
        return
    progress_var.set(0)
    log_queue.put("Starting crawl...")
    threading.Thread(target=crawl_websites, args=(input_file, output_file, progress_var, log_queue, total_websites)).start()

def update_log():
    try:
        while True:
            log_message = log_queue.get_nowait()
            log_text.insert(tk.END, log_message + "\n")
            log_text.see(tk.END)
    except queue.Empty:
        pass
    root.after(100, update_log)

root = tk.Tk()
root.title("Indian Phone Crawler")

input_label = tk.Label(root, text="Input Excel File:")
input_label.grid(row=0, column=0, sticky="w", padx=5, pady=5)
input_entry = tk.Entry(root, width=50)
input_entry.grid(row=0, column=1, padx=5, pady=5)
input_button = tk.Button(root, text="Browse", command=browse_input_file)
input_button.grid(row=0, column=2, padx=5, pady=5)

output_label = tk.Label(root, text="Output Excel File:")
output_label.grid(row=1, column=0, sticky="w", padx=5, pady=5)
output_entry = tk.Entry(root, width=50)
output_entry.grid(row=1, column=1, padx=5, pady=5)
output_button = tk.Button(root, text="Browse", command=browse_output_file)
output_button.grid(row=1, column=2, padx=5, pady=5)

progress_label = tk.Label(root, text="Progress:")
progress_label.grid(row=2, column=0, sticky="w", padx=5, pady=5)
progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100)
progress_bar.grid(row=2, column=1, columnspan=2, sticky="ew", padx=5, pady=5)
progress_bar["style"] = "green.Horizontal.TProgressbar"
style = ttk.Style(root)
style.configure("green.Horizontal.TProgressbar", troughcolor='#d9d9d9', background='green')

start_button = tk.Button(root, text="Start Crawl", command=start_crawl)
start_button.grid(row=3, column=1, pady=10)

log_label = tk.Label(root, text="Progress Log:")
log_label.grid(row=4, column=0, sticky="w", padx=5, pady=5)
log_text = tk.Text(root, height=10, width=60)
log_text.grid(row=5, column=0, columnspan=3, padx=5, pady=5)
log_scrollbar = tk.Scrollbar(root, command=log_text.yview)
log_scrollbar.grid(row=5, column=3, sticky="ns")
log_text.config(yscrollcommand=log_scrollbar.set)

log_queue = queue.Queue()
update_log()

root.mainloop()