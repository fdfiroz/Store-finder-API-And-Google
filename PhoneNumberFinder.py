import time
import urllib.parse  # Added for URL encoding
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

def init_driver():
    # Set up Chrome options
    chrome_options = Options()
    # Uncomment the next line to run headless mode if needed:
    # chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    
    # If you need to load a CRX extension, uncomment below and update the path:
    chrome_options.add_extension("C:/chromedriver/NopeCHA-CAPTCHA-Solver-Chrome-Web-Store.crx")
    
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.maximize_window()
    return driver

def extract_info(driver, store_name, address, city, state, postal_code):
    """
    Constructs a Google search query using the store name and address details,
    navigates to the search results page, and extracts the phone number,
    website URL, and closed status.
    """
    # Build the query string; note that if store_name contains "&", it needs encoding.
    query = f"{store_name}, {address}, {city}, {state} {postal_code} phone"
    # URL-encode the entire query so special characters like & are preserved.
    encoded_query = urllib.parse.quote_plus(query)
    search_url = "https://www.google.com/search?q=" + encoded_query + "&hl=en"
    print(f"\nSearching for: {query}")
    print(f"Encoded Search URL: {search_url}")
    
    driver.get(search_url)
    time.sleep(3)  # Allow time for the page to load
    
    # Check for CAPTCHA or robot verification (if present, wait for manual intervention)
    if "unusual traffic" in driver.title.lower():
        print("[!] CAPTCHA detected. Please solve it manually and press Enter...")
        input()
    
    phone = None
    website = None
    is_closed = False

    # --- Extract Phone Number ---
    try:
        phone_element = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located(
                (By.XPATH, "//span[contains(@aria-label, 'Call phone number')]")
            )
        )
        aria_label = phone_element.get_attribute("aria-label")
        print("Found phone aria-label:", aria_label)
        if aria_label.startswith("Call phone number "):
            phone = aria_label[len("Call phone number "):].strip()
        else:
            phone = aria_label.strip()
    except Exception as e:
        print(f"[!] Phone extraction error for query '{query}': {e}")

    # --- Extract Website URL ---
    try:
        website_element = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located(
                (By.XPATH, "//div[contains(@class, 'IzNS7c')]//a[contains(@class, 'ab_button') and .//div[text()='Website']]")
            )
        )
        website = website_element.get_attribute("href")
        print("Found website:", website)
    except Exception as e:
        print(f"[!] Website extraction error for query '{query}': {e}")

    # --- Check Closed Status (Permanently or Temporarily closed) ---
    try:
        # Look for "Permanently closed"
        closed_element = driver.find_element(By.XPATH, "//*[contains(text(), 'Permanently closed')]")
        if closed_element:
            is_closed = True
            print("Business is marked as Permanently closed.")
    except Exception:
        try:
            # Look for "Temporarily closed"
            closed_element = driver.find_element(By.XPATH, "//*[contains(text(), 'Temporarily closed')]")
            if closed_element:
                is_closed = True
                print("Business is marked as Temporarily closed.")
        except Exception:
            is_closed = False

    return phone, website, is_closed

def main():
    driver = init_driver()
    
    # Input Excel file must have columns: Store Name, Address, City, State, PostalCode
    input_file = "stores.xlsx"
    output_file = "stores_with_info.xlsx"
    df = pd.read_excel(input_file)
    
    phones = []
    websites = []
    closed_status = []
    
    for idx, row in df.iterrows():
        store_name = row["Store Name"]
        address = row["Address"]
        city = row["City"]
        state = row["State"]
        postal_code = str(row["PostalCode"])
        
        print(f"\nProcessing: {store_name}, {address}, {city}, {state} {postal_code}")
        phone, website, is_closed = extract_info(driver, store_name, address, city, state, postal_code)
        print(f"--> Phone: {phone}")
        print(f"--> Website: {website}")
        print(f"--> Closed status: {is_closed}")
        
        phones.append(phone)
        websites.append(website)
        closed_status.append(is_closed)
        
        time.sleep(2)  # Delay between queries
    
    df["Phone Number"] = phones
    df["Website"] = websites
    df["Closed"] = closed_status
    df.to_excel(output_file, index=False)
    print(f"\nResults saved to {output_file}")
    
    driver.quit()

if __name__ == "__main__":
    main()
