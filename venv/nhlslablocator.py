import pandas as pd
import googlemaps
import openpyxl

# Your API Key
gmaps = googlemaps.Client(key='AIzaSyABAPas8lhxjnKzVKqtxHclUQDM-VsXSEM')
# Function to handle geocoding
def geocode_address(address):
    try:
        # Geocode the address
        geocode_result = gmaps.geocode(address, region='za')

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

# Create new 'latitude', 'longitude', and 'status' columns
df['latitude'], df['longitude'], df['status'] = zip(*df['CTUSL_Desc'].astype(str).apply(lambda x: geocode_address(x.replace('|', ', '))))

# Save the data back to .xlsx
df.to_excel(r"C:\Users\cafra\Desktop\User site addresses.xlsx", index=False, engine='openpyxl')
