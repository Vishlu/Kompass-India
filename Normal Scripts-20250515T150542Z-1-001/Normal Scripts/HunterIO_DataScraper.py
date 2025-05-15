from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
import pandas as pd
import time
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

def initialize_driver():
    """Initialize the WebDriver and return the driver instance."""
    try:
        driver = webdriver.Chrome()
        logging.info("WebDriver initialized successfully.")
        return driver
    except Exception as e:
        logging.error(f"Error initializing WebDriver: {e}")
        raise

def open_hunter_login(driver):
    """Navigate to the Hunter.io login page."""
    try:
        driver.get("https://hunter.io/discover?filters_changed=1&industry_excluded=&industry_included=54&location_business_region_excluded=&location_business_region_included=&location_city_excluded=&location_city_included=&location_continent_excluded=&location_continent_included=&location_country_excluded=&location_country_included=IN&location_state_excluded=&location_state_included=&tab=technologies")
        time.sleep(2)
        logging.info("Navigated to Hunter.io login page.")
    except Exception as e:
        logging.error(f"Error navigating to Hunter.io login page: {e}")
        raise

def click_google_sign_in(driver):
    """Click on the 'Sign in with Google' button."""
    try:
        google_sign_in_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "btn-lg.btn-google.full-width"))
        )
        google_sign_in_button.click()
        time.sleep(3)
        logging.info("Clicked on 'Sign in with Google' button.")
    except Exception as e:
        logging.error(f"Error clicking 'Sign in with Google' button: {e}")
        raise

def enter_email(driver, email):
    """Enter the email address in the Google login page."""
    try:
        email_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "identifierId"))
        )
        email_input.send_keys(email)
        email_input.send_keys(Keys.RETURN)
        time.sleep(2)
        logging.info("Entered email address.")
    except Exception as e:
        logging.error(f"Error entering email address: {e}")
        raise

def enter_password(driver, password):
    """Enter the password in the Google login page."""
    try:
        password_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "Passwd"))
        )
        password_input.send_keys(password)
        password_input.send_keys(Keys.RETURN)
        time.sleep(5)
        logging.info("Entered password.")
    except Exception as e:
        logging.error(f"Error entering password: {e}")
        raise
    


def click_email_addresses_tab(driver):
    """Clicks the 'Email addresses' tab."""
    try:
        # Click on the first tab (Email addresses)
        email_tab_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Email addresses')]"))
        )
        email_tab_button.click()
        logging.info("Clicked on 'Email addresses' tab.")

        # Wait for the email section to load
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "company-details-emails"))
        )
        logging.info("Email addresses tab content loaded.")
    except Exception as e:
        logging.error(f"Error clicking 'Email addresses' tab: {e}")
        raise

def extract_executive_names_and_designations(driver):
    """Extracts executive names and designations from the email addresses tab."""
    extracted_data = []
    try:
        soup = BeautifulSoup(driver.page_source, 'html.parser')

        # Extract all the executive data in the "ds-result" class
        results = soup.find_all('div', class_='ds-result__data')

        for result in results:
            name_tag = result.find('div', class_='ds-result__fullname')
            designation_tag = result.find('div', class_='ds-result__secondary')

            # Extract name
            if name_tag:
                executive_name = name_tag.get_text(strip=True)
            else:
                executive_name = None

            # Extract designation
            if designation_tag:
                designation_tag = designation_tag.find_all('div', class_='ds-result__attribute')
                designation = ", ".join([tag.get_text(strip=True) for tag in designation_tag if tag.get_text(strip=True)]) if designation_tag else None
            else:
                designation = None

            # Check if both executive name and designation are found
            if executive_name and designation:
                extracted_data.append(f"{executive_name} - {designation}")
            else:
                extracted_data.append("")  # Append empty if either is missing

        if extracted_data:
            return ", ".join([data for data in extracted_data if data])  # Return data, excluding empty results
        else:
            logging.warning("No executive data found.")
            return ""

    except Exception as e:
        logging.error(f"Error extracting executive data: {e}")
        return ""


# def click_technologies_tab(driver):
#     try:
#         # More robust way to find the Technologies tab.
#         technologies_tab_button = WebDriverWait(driver, 20).until( #increased timeout
#             EC.element_to_be_clickable((By.XPATH, "//button[text()='Technologies']")) 
#         )
#         technologies_tab_button.click()
#         logging.info("Technologies tab clicked successfully.")

#         # Important: Wait for the technologies content to load dynamically.
#         WebDriverWait(driver, 20).until( # Increased timeout and more specific condition
#             EC.presence_of_element_located((By.ID, "discover_technologies")) #wait for turbo-frame
#         )
#         logging.info("Technologies content loaded.")


#     except Exception as e:
#         logging.error(f"Error occurred while clicking the Technologies tab: {str(e)}")
#         return False # Indicate failure
#     return True # Indicate Success




# def extract_technologies_data(driver):
#     try:
#         # Wait for the technologies section to be fully loaded
#         WebDriverWait(driver, 10).until(
#             EC.presence_of_element_located((By.CSS_SELECTOR, '.key-value__item'))
#         )

#         # Get the page source after the dynamic content is fully loaded
#         soup = BeautifulSoup(driver.page_source, 'html.parser')

#         tech_data = {}

#         # Find all technology items
#         tech_items = soup.find_all('div', class_='key-value__item')

#         if not tech_items:
#             logging.info("No technology items found on the page.")
#             print("Page Source (Technology Items Check):", driver.page_source)  # Debugging message
#             return None  # Return None if no technology items are found

#         for item in tech_items:
#             # Extract technology name
#             tech_name_tag = item.find('dt', class_='key-value__title')
#             tech_name = tech_name_tag.text.strip() if tech_name_tag else "N/A"

#             # Extract tools associated with the technology
#             tool_names = []
#             dd_tags = item.find_all('dd', class_='key-value__desc')

#             for dd_tag in dd_tags:
#                 tool_span = dd_tag.find('span')
#                 tool_name = tool_span.text.strip() if tool_span else "N/A"
#                 tool_names.append(tool_name)

#             # Add the technology data to the dictionary
#             tech_data[tech_name] = ", ".join(tool_names)

#         # Format the extracted data for logging or further use
#         formatted_data = " / ".join([f"{tech}: {tools}" for tech, tools in tech_data.items()])
#         logging.info(f"Extracted Technology Data: {formatted_data}")

#         return formatted_data

#     except Exception as e:
#         logging.error(f"Error occurred while extracting technology data: {str(e)}")
#         return None



def extract_company_data(driver):
    """Extracts company details from the opened company page using BeautifulSoup and additional tabs."""
    try:
        soup = BeautifulSoup(driver.page_source, 'html.parser')

        # Extract Company Name
        company_name_tag = soup.find('h2', class_='company-details__name')
        company_name = company_name_tag.get_text(strip=True) if company_name_tag else ""

        # Extract Description
        description_tag = soup.find('div', class_='company-details__description')
        description = description_tag.get_text(strip=True) if description_tag else ""

        # Extract Key-Value Fields
        industry = size = country = year_founded = company_type = website = keywords = ""

        key_value_items = soup.find_all('div', class_='key-value__item')

        for item in key_value_items:
            title_tag = item.find('dt', class_='key-value__title')
            desc_tag = item.find('dd', class_='key-value__desc')

            if title_tag and desc_tag:
                title = title_tag.get_text(strip=True)
                desc = desc_tag.get_text(strip=True)

                if title == "Industry:":
                    industry = desc
                elif title == "Size:":
                    size = desc
                elif title == "Address:":
                    country = desc
                elif title == "Year founded:":
                    year_founded = desc
                elif title == "Type:":
                    company_type = desc
                elif title == "Website:":
                    website_tag = desc_tag.find('a')
                    website = website_tag['href'] if website_tag else ""
                elif title == "Keywords:":
                    keywords = desc

        # Extract Social Links
        social_links = []
        social_tags = soup.select('.company-details__socials a')

        for tag in social_tags:
            social_links.append(tag['href'])

        social_links = ", ".join(social_links) if social_links else ""


        return {
            "Company Name": company_name,
            "Description": description,
            "Industry": industry,
            "Size": size,
            "Country": country,
            "Year Founded": year_founded,
            "Type": company_type,
            "Website": website,
            "Keywords": keywords,
            "Social": social_links,
        }

    except Exception as e:
        logging.error(f"Error extracting company data: {e}")
        return {}




def click_all_company_rows(driver):
    """Clicks all rows of the company table one by one, extracts data, closes the details view, and moves to the next one."""
    extracted_data = []
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "tbody tr.h-table__row-link"))
        )

        clicked_rows = set()

        while True:
            rows = driver.find_elements(By.CSS_SELECTOR, "tbody tr.h-table__row-link")
            if not rows:
                logging.info("No company rows found.")
                break  

            for row in rows:
                row_id = row.get_attribute("data-address-bar-url-param")

                if row_id not in clicked_rows:
                    try:
                        driver.execute_script("arguments[0].scrollIntoView();", row)
                        driver.execute_script("arguments[0].click();", row)
                        clicked_rows.add(row_id)

                        logging.info(f"Clicked on row: {row_id}")

                        # Wait for company page to load
                        time.sleep(3)

                        # Extract Data
                        company_data = extract_company_data(driver)
                        
                         # Click the 'Email addresses' tab before extracting executive data
                        click_email_addresses_tab(driver)
                        
                        # Extract Executive Data (Names and Designations)
                        executive_data = extract_executive_names_and_designations(driver)
                        
                        # # Step 1: Click on the Technologies tab
                        # click_technologies_tab(driver)

                        
                        
                        #  # Click Technologies and Extract (Improved)
                        # if click_technologies_tab(driver): #Check if tab click was successful
                        #     technologies_data = extract_technologies_data(driver)
                        if company_data:
                            company_data["Executives and Designation"] = executive_data
                                # company_data["Technologies"] = technologies_data
                            extracted_data.append(company_data)
                        # else:
                        #     logging.warning("Failed to click Technologies tab for this company.")

                        # Close the company details by clicking the close button
                        close_button = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[aria-label="Close company details"]'))
                        )
                        close_button.click()
                        logging.info("Closed company details.")

                        # Wait for the page to load back to the list
                        time.sleep(3)

                    except Exception as e:
                        logging.warning(f"Could not click row: {e}")
                        continue

    except Exception as e:
        logging.error(f"Error clicking all company rows: {e}")
    
    return extracted_data

def save_to_excel(data, filename=r"C:\Users\Kompass\Desktop\Hunter IO Files\Chemical_Manufacturing_Hunter_io_Data_4-2-2025.xlsx"):
    """Saves extracted company data into an Excel file."""
    try:
        df = pd.DataFrame(data)
        df.to_excel(filename, index=False)
        logging.info(f"Data saved to {filename} successfully.")
    except Exception as e:
        logging.error(f"Error saving to Excel: {e}")

def main():
    """Main function to execute the login and data extraction process."""
    driver = None
    try:
        driver = initialize_driver()
        open_hunter_login(driver)
        click_google_sign_in(driver)
        enter_email(driver, "datascience@kompassindia.com")
        enter_password(driver, "Kompass@India@DataScience@2025")

        time.sleep(6)

        # Extract company data
        extracted_data = click_all_company_rows(driver)

        # Save data to Excel
        save_to_excel(extracted_data)

        logging.info("Data extraction and saving to Excel completed.")

    except Exception as e:
        logging.error(f"An error occurred during execution: {e}")

    finally:
        if driver:
            driver.quit()
            logging.info("WebDriver closed successfully.")

if __name__ == "__main__":
    main()
