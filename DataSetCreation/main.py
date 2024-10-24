import pandas as pd
import requests
import re
from datetime import datetime

file_path = 'add_your_excel_file_path'
df = pd.read_excel(file_path)

# VisualCrossing API details
VISUALCROSSING_API_KEY = 'add_your_api_key'
VISUALCROSSING_BASE_URL = 'https://weather.visualcrossing.com/VisualCrossingWebServices/rest/services/timeline/'
GOOGLE_MAPS_API_KEY = 'add_your_api_key'
GOOGLE_MAPS_BASE_URL = 'https://maps.googleapis.com/maps/api/geocode/json'


def clean_address(address):
    # Remove everything except letters and commas; focus on city/place names
    cleaned_address = re.sub(r'NO\.?\s*|NO\s*/?\s*|[0-9]+[.,]*|/+', '', address)  # Remove numbers, NO., and slashes
    cleaned_address = re.sub(r'\b\d{9,}\b', '', cleaned_address)  # Remove phone numbers (9+ digits)
    cleaned_address = re.sub(r'\bROAD\b|\bRD\b|\bHILL\b|\bDIVISION\b|\bESTATE\b', '', cleaned_address)  # Remove unnecessary words
    cleaned_address = re.sub(r'\s+', ' ', cleaned_address)  # Remove excessive whitespace
    cleaned_address = cleaned_address.strip()  # Remove leading/trailing whitespace
    return cleaned_address


# Function to get latitude and longitude from Google Maps Geocoding API
def get_lat_lon_from_google(address):
    google_maps_url = f"{GOOGLE_MAPS_BASE_URL}?address={address}&key={GOOGLE_MAPS_API_KEY}"
    response = requests.get(google_maps_url)

    if response.status_code == 200:
        geocode_data = response.json()
        if geocode_data['results']:
            location = geocode_data['results'][0]['geometry']['location']
            lat = location['lat']
            lon = location['lng']
            return lat, lon
        else:
            print(f"Failed to get geocode data for {address}. No results found.")
            return None, None
    else:
        print(f"Failed to get geocode data for {address}. Status code: {response.status_code}")
        return None, None


# Function to get weather data from VisualCrossing API using lat/lon
def get_weather_data(lat, lon, date_time):
    formatted_date_time = date_time.strftime('%Y-%m-%dT%H:%M:%S')
    api_url = f"{VISUALCROSSING_BASE_URL}{lat},{lon}/{formatted_date_time}?key={VISUALCROSSING_API_KEY}&unitGroup=metric&include=current"

    response = requests.get(api_url)

    if response.status_code == 200:
        return response.json()
    else:
        try:
            error_message = response.json().get('message', 'No specific error message provided.')
        except Exception:
            error_message = "Failed to parse error message."
        print(f"Failed to get data for coordinates ({lat}, {lon}) at {formatted_date_time}. "
              f"Status code: {response.status_code}. Reason: {error_message}")
        return None


df['Temperature'] = None
df['Humidity'] = None
df['WindSpeed'] = None
df['Conditions'] = None
df['Precipitation'] = None
df['PrecipitationType'] = None


for index, row in df.iterrows():
    received_time = pd.to_datetime(row['Received Time'])
    address = row['Address']

    # Clean and extract the relevant part of the address
    cleaned_address = clean_address(address)

    # Get latitude and longitude from Google Maps API
    lat, lon = get_lat_lon_from_google(cleaned_address)

    if lat and lon:
        # Get weather data from VisualCrossing API
        weather_data = get_weather_data(lat, lon, received_time)

        if weather_data:
            current_conditions = weather_data.get('currentConditions', {})
            df.at[index, 'Temperature'] = current_conditions.get('temp', 'N/A')
            df.at[index, 'Humidity'] = current_conditions.get('humidity', 'N/A')
            df.at[index, 'WindSpeed'] = current_conditions.get('windspeed', 'N/A')
            df.at[index, 'Conditions'] = current_conditions.get('conditions', 'N/A')
            df.at[index, 'Precipitation'] = current_conditions.get('precip', 'N/A')
            df.at[index, 'PrecipitationType'] = current_conditions.get('preciptype', 'N/A')


output_file = 'add_the_location_to_save_excel_file'
df.to_excel(output_file, index=False)
print(f"Updated file saved to {output_file}")
