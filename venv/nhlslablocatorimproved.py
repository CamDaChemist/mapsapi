import os
import pandas as pd
import googlemaps
import openpyxl

# Your API Key
gmaps = googlemaps.Client(key='AIzaSyABAPas8lhxjnKzVKqtxHclUQDM-VsXSEM')

import os
import pandas as pd
import googlemaps
import openpyxl


# Function to handle geocoding
def geocode_address_single(address):
    try:
        # Geocode the address
        geocode_result = gmaps.geocode(address.replace('|', ', '), region='za')

        # If result is empty, return None for coordinates and "Not Found" status
        if not geocode_result:
            return None, None, "Not Found"

        # Extract coordinates from geocode result
        lat = geocode_result[0]['geometry']['location']['lat']
        lng = geocode_result[0]['geometry']['location']['lng']

        # Return coordinates and "Found" status
        return lat, lng, "Found"

    except Exception as e:
        # Return None for coordinates and the error message as status
        return None, None, str(e)

# Load the data
df = pd.read_excel(r"C:\Users\cafra\Desktop\User site addresses.xlsx", engine='openpyxl')

# Geocode for each column and create new 'latitude', 'longitude', and 'status' columns for each
for column in ['CTUSL_Address', 'CTUSL_Address1', 'CTUSL_Address2', 'CTUSL_Address3_Suburb']:
    df[column + '_lat'], df[column + '_lng'], df[column + '_status'] = zip(*df[column].astype(str).apply(lambda x: geocode_address_single(x)))

# Save the data back to .xlsx
df.to_excel(r"C:\Users\cafra\Desktop\User site addresses.xlsx", index=False, engine='openpyxl')
