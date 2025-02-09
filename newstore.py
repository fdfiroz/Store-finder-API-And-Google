import pandas as pd
import requests
import time
import random
import re

def parse_address(address):
    """Extract city, state, and zip code from address string"""
    try:
        # Use regex to find state (2 letters) and zip (5 digits) at the end
        match = re.search(r'([A-Z]{2})\s(\d{5})$', address)
        if match:
            state, zip_code = match.groups()
            city_part = address.rsplit(f' {state} {zip_code}', 1)[0]
            # Find city as the last word(s) before state
            city = city_part.split()[-1]
            if city.isupper():  # Handle cases where city might be abbreviated
                city = city.title()
            return city, state, zip_code
        return '', '', ''
    except:
        return '', '', ''

# Read input Excel file with coordinates
input_file = 'uscities.xlsx'
df = pd.read_excel(input_file)

# Verify required columns
if 'latitude' not in df.columns or 'longitude' not in df.columns:
    raise ValueError("Excel file must contain 'latitude' and 'longitude' columns")

results = []

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36',
    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'Origin': 'https://www.khallstudio.com',
    'Referer': 'https://www.khallstudio.com/locator/',
    'X-Requested-With': 'XMLHttpRequest'
}

for index, row in df.iterrows():
    lat = row['latitude']
    lon = row['longitude']
    
    payload = {
        'action': 'acf_locations_limit',
        'type': 'retailer',
        'distance': '500',
        'latitude': lat,
        'longitude': lon
    }
    
    try:
        # Add random delay between requests (2-15 seconds)
        delay = random.uniform(3, 10)
        time.sleep(delay)
        
        response = requests.post(
            'https://www.khallstudio.com/wp-admin/admin-ajax.php',
            headers=headers,
            data=payload,
            timeout=15
        )
        response.raise_for_status()
        
        locations = response.json()
        
        for location in locations:
            city, state, zip_code = parse_address(location.get('address', ''))
            
            entry = {
                # 'query_latitude': lat,
                # 'query_longitude': lon,
                # 'result_latitude': location.get('latitude'),
                # 'result_longitude': location.get('longitude'),
                'name': location.get('name', ''),
                'address': location.get('address', ''),
                'city': city,
                'state': state,
                "mapaddress": location.get('mapaddress', ''),
                'postal_code': zip_code,
                'country': 'USA',
                'phone': location.get('phone', '').replace('-', ''),
                'distance': location.get('distance', ''),
                'raw_data': str(location)  # Store complete data for reference
            }
            results.append(entry)
            
        print(f"Processed {len(locations)} locations for coordinates {lat},{lon}")
        
    except Exception as e:
        print(f"Error processing {lat},{lon}: {str(e)}")

# Save results to Excel
if results:
    output_df = pd.DataFrame(results)
    output_file = 'khall_locations.xlsx'
    
    # Custom column order
    columns = [
        
        'name', 'address', 'city', 'state', 'postal_code', 'country',
        'phone', 'distance', 'mapaddress'
    ]
    
    output_df[columns].to_excel(output_file, index=False)
    print(f"Saved {len(results)} locations to {output_file}")
else:
    print("No locations found")