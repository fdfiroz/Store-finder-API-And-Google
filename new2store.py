import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options

def scrape_spas(city_file, output_file):
    cities_df = pd.read_excel(city_file)
    cities = cities_df['City'].tolist()

    try:
        chrome_options = Options()
        chrome_options.add_argument("--headless")  # run browser in the background
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.maximize_window()
    except Exception as e:
        print("\nChrome WebDriver Error: Please follow these steps:")
        print("1. Install webdriver_manager using: pip install webdriver_manager")
        print("2. Make sure you have Chrome browser installed")
        print(f"Error details: {str(e)}")
        raise e

    url = 'https://farmhousefreshgoods.com/pages/find-a-spa'
    driver.get(url)
    time.sleep(5)  # Allow page to fully load

    # ✅ FIX: Close popup if it exists
    try:
        popup_close = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".klaviyo-close-form, .popup-close, .close-button"))
        )
        popup_close.click()
        print("✅ Popup closed successfully!")
        time.sleep(2)
    except:
        print("⚠ No popup found or already closed")

    # ✅ FIX: Wait for search elements dynamically
    try:
        search_box = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, 'storemapper-zip'))
        )
        search_button = driver.find_element(By.ID, 'storemapper-go')

        # ✅ FIX: Ensure dropdown button is interactable
        radius_dropdown_button = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.ID, 'storemapper-distance-btn'))
        )

        print("✅ Search elements found successfully!")
    except Exception as e:
        print("❌ Error locating search elements:", e)
        driver.quit()
        return

    # Create an empty Excel file before starting to store data
    df_empty = pd.DataFrame(columns=['City', 'Spa Name', 'Address', 'Phone', 'Distance', 'Directions URL', 'Description', 'Services'])
    df_empty.to_excel(output_file, index=False)

    for city in cities:
        print(f"🔎 Searching for spas in: {city}")

        # Enter new city in the search box
        search_box.clear()
        search_box.send_keys(city)
        time.sleep(1)

        # ✅ FIX: Scroll down for better visibility
        driver.execute_script("window.scrollBy(0, 250);")
        time.sleep(1)

        # ✅ FIX: Scroll the dropdown into view before clicking
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", radius_dropdown_button)
        time.sleep(1)

        # ✅ FIX: Click dropdown button
        try:
            radius_dropdown_button.click()
            time.sleep(1)
        except:
            driver.execute_script("arguments[0].click();", radius_dropdown_button)
            time.sleep(1)

        # ✅ FIX: Click the 250 mi radio button
        try:
            max_radius_option = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID, 'storemapperRadius-250'))
            )
            max_radius_option.click()
            time.sleep(1)
        except:
            print("❌ Failed to select 250 mi radius")

        # Click search button
        search_button.click()
        time.sleep(3)

        # Wait for results to load
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'storemapper-list'))
            )
        except:
            print(f"❌ No results found for {city}")
            continue

        # ✅ FIX: Scroll inside the search results container for loading more stores
        try:
            results_container = driver.find_element(By.ID, 'storemapper-list')
        except:
            print(f"❌ Could not locate results container for {city}")
            continue

        # ✅ FIX: Click "Show More Stores" until all results load, scrolling inside the results section
        while True:
            try:
                show_more_button = WebDriverWait(driver, 3).until(
                    EC.element_to_be_clickable((By.CLASS_NAME, 'strmpr-view-more-stores-button'))
                )

                driver.execute_script("arguments[0].scrollBy(0, 300);", results_container)
                time.sleep(1)

                if 'strmpr-hidden' in show_more_button.get_attribute('class'):
                    break  # No more stores to load

                show_more_button.click()
                time.sleep(2)

            except:
                break  # No "Show More" button found, move on

        # ✅ Extract spa details
        spa_data = []
        listings = driver.find_elements(By.CSS_SELECTOR, '.strmpr-search-result')
        for listing in listings:
            try:
                name = listing.find_element(By.CSS_SELECTOR, '.strmpr-field-name').text
                address = listing.find_element(By.CSS_SELECTOR, '.strmpr-field-address').text
                phone = listing.find_element(By.CSS_SELECTOR, '.strmpr-field-phone a').text if listing.find_elements(By.CSS_SELECTOR, '.strmpr-field-phone a') else "N/A"
                distance = listing.find_element(By.CSS_SELECTOR, '.strmpr-field-distance').text if listing.find_elements(By.CSS_SELECTOR, '.strmpr-field-distance') else "N/A"
                directions_url = listing.find_element(By.CSS_SELECTOR, '.strmpr-field-directions a').get_attribute("href") if listing.find_elements(By.CSS_SELECTOR, '.strmpr-field-directions a') else "N/A"
                description = listing.find_element(By.CSS_SELECTOR, '.strmpr-field-description').text if listing.find_elements(By.CSS_SELECTOR, '.strmpr-field-description') else "N/A"

                services = []
                custom_fields = listing.find_elements(By.CSS_SELECTOR, '.strmpr-field-custom')
                for field in custom_fields:
                    service_text = field.text.strip()
                    if service_text:
                        services.append(service_text)

                spa_data.append({
                    'City': city,
                    'Spa Name': name,
                    'Address': address,
                    'Phone': phone,
                    'Distance': distance,
                    'Directions URL': directions_url,
                    'Description': description,
                    'Services': "; ".join(services)  # Join services into one field
                })
            except Exception as e:
                print(f"⚠ Error extracting data for a listing in {city}: {e}")
                continue

        # ✅ Save city data immediately
        if spa_data:
            df_city = pd.DataFrame(spa_data)
            with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                df_city.to_excel(writer, index=False, header=False, startrow=writer.sheets['Sheet1'].max_row)

            print(f"✅ Saved results for {city} immediately!")

    driver.quit()
    print(f"✅ Scraping completed! All data saved to {output_file}")

# Main execution
if __name__ == "__main__": 
    input_file = 'uscities.xlsx'
    output_file = 'spa_locations.xlsx'
    scrape_spas(input_file, output_file)
