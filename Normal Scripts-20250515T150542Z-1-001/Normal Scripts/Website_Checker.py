import requests

def check_website_status(url):
    try:
        response = requests.get(url)
        # Check if the status code is 200 (OK)
        if response.status_code == 200:
            return f"The website {url} is active (Status Code: {response.status_code})."
        else:
            return f"The website {url} is inactive (Status Code: {response.status_code})."
    except requests.exceptions.RequestException as e:
        return f"The website {url} is inactive. Error: {e}"

# Example usage
url = "http://kompassxyz.com/"
status = check_website_status(url)
print(status)