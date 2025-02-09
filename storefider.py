import pandas as pd
import requests
import json
from time import sleep

def fetch_store_data(latitude, longitude):
    # API endpoint
    url = f"https://stockist.co/api/v1/u2517/locations/search?tag=u2517&latitude={latitude}&longitude={longitude}"
    
    try:
        response = requests.get(url)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"Error fetching data: {e}")
        return None

def extract_store_info(location):
    return {
        'name': location.get('name'),
        'address_line_1': location.get('address_line_1'),
        'address_line_2': location.get('address_line_2'),
        'city': location.get('city'),
        'state': location.get('state'),
        'postal_code': location.get('postal_code'),
        'country': location.get('country'),
        'phone': location.get('phone'),
        'website': location.get('website')
    }

def main():
    # Read coordinates from Excel
    try:
        df_coordinates = pd.read_excel('uscities.xlsx')
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    # Store all results
    all_stores = []

    # Process each coordinate
    for index, row in df_coordinates.iterrows():
        lat = row['lat']
        lng = row['lng']
        
        # Fetch data
        data = fetch_store_data(lat, lng)
        
        if data and 'locations' in data:
            for location in data['locations']:
                store_info = extract_store_info(location)
                all_stores.append(store_info)
        
        # Add delay to avoid rate limiting
        sleep(1)

    # Save results to Excel
    if all_stores:
        df_results = pd.DataFrame(all_stores)
        df_results.to_excel('store_results.xlsx', index=False)
        print("Results saved to store_results.xlsx")
    else:
        print("No data found")

if __name__ == "__main__":
    main()