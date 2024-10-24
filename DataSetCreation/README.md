# Weather Data Augmentation Script

This Python script retrieves weather data for a set of locations and timestamps listed in an Excel file. It uses the Google Maps Geocoding API to convert addresses into latitude and longitude, and then the VisualCrossing API to fetch weather conditions for each location and time.

## Features

- Cleans the address data to ensure accurate geolocation.
- Retrieves latitude and longitude for each location using the Google Maps Geocoding API.
- Fetches weather data (temperature, humidity, wind speed, conditions, precipitation, and precipitation type) from the VisualCrossing API.
- Adds the retrieved weather data to the original Excel file.

## Requirements

- Python 3.x
- Required Libraries:
  - `pandas`
  - `requests`
  - `re` (built-in)
  - `datetime` (built-in)

You can install the necessary libraries by running:

```bash
pip install pandas requests
