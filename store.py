import pandas as pd
import requests
import time
import random

# Read the input Excel file
input_file = 'uscities.xlsx'  # Update with your input file name
df = pd.read_excel(input_file)

# Check if required columns exist
if 'latitude' not in df.columns or 'longitude' not in df.columns:
    raise ValueError("Excel file must contain 'latitude' and 'longitude' columns.")

results = []
total_requests = len(df)
completed_requests = 0

for index, row in df.iterrows():
    lat = row['latitude']
    lon = row['longitude']
    
    # Construct the API URL with current latitude and longitude
    url = f"https://stockist.co/api/v1/u13410/locations/search?tag=u2517&latitude={lat}&longitude={lon}"
    
    try:
        # Send GET request to the API
        response = requests.get(url, timeout=10)
        response.raise_for_status()  # Raise exception for HTTP errors
        completed_requests += 1
        print(f"Completed request {completed_requests} of {total_requests} ({(completed_requests/total_requests)*100:.1f}%)")
        
        # Parse JSON response
        data = response.json()
        # print(data)
        
        # Extract required fields from each location
        for location in data.get('locations', []):
            entry = {
                'query_latitude': lat,
                'query_longitude': lon,
                'name': location.get('name', ''),
                'address_line_1': location.get('address_line_1', ''),
                'address_line_2': location.get('address_line_2', '') or '',
                'city': location.get('city', ''),
                'state': location.get('state', ''),
                'postal_code': location.get('postal_code', ''),
                'country': location.get('country', ''),
                'phone': location.get('phone', ''),
                'website': location.get('website', '')
            }
            results.append(entry)
    
    except requests.exceptions.RequestException as e:
        print(f"Request failed for lat={lat}, lon={lon}: {e}")
    except ValueError as e:
        print(f"Failed to parse JSON response for lat={lat}, lon={lon}: {e}")
    
    
    start_time = time.time()

    delay = random.uniform(2, 6)
    # print(f"Waiting for {delay:.2f} seconds before next request...")
    time.sleep(delay)

# Create a DataFrame from the results and save to Excel
if results:
    results_df = pd.DataFrame(results)
    output_file = 'output.xlsx'  # Update with your desired output file name
    results_df.to_excel(output_file, index=False)
    print(f"Data successfully saved to {output_file}")
    end_time = time.time()
    total_time = end_time - start_time
    # print(f"\nTotal scraping time: {total_time:.2f} seconds ({total_time/60:.2f} minutes)")
else:
    print("No data was fetched from the API.")