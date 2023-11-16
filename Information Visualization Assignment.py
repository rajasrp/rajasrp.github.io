#!/usr/bin/env python
# coding: utf-8

# In[3]:


import pandas as pd
import requests

def get_geolocation(api_key, location):
    base_url = 'https://maps.googleapis.com/maps/api/geocode/json'
    params = {
        'address': location,
        'key': api_key,
    }

    response = requests.get(base_url, params=params)
    data = response.json()

    if data['status'] == 'OK':
        location = data['results'][0]['geometry']['location']
        latitude = location['lat']
        longitude = location['lng']
        return latitude, longitude
    else:
        print(f"Error: {data['status']}")
        return None

api_key = 'API_KEY'

file_path = 'Copy of Bioscience Map of Maine company-town (2).xlsx'  
sheet_name = 'Colleges'  
column_name = 'Colleges'  

df = pd.read_excel(file_path, sheet_name=sheet_name)

for index, row in df.iterrows():
    institution_name = row[column_name]
    location = get_geolocation(api_key, institution_name)

    if location:
        print(f"{institution_name}: Latitude - {location[0]}, Longitude - {location[1]}")


# In[4]:


import pandas as pd
import requests

def get_geolocation(api_key, location):
    base_url = 'https://maps.googleapis.com/maps/api/geocode/json'
    params = {
        'address': location,
        'key': api_key,
    }
    response = requests.get(base_url, params=params)
    data = response.json()

    if data['status'] == 'OK':
        location = data['results'][0]['geometry']['location']
        latitude = location['lat']
        longitude = location['lng']
        return latitude, longitude
    else:
        print(f"Error: {data['status']}")
        return None


api_key = 'API_KEY'

file_path = 'Copy of Bioscience Map of Maine company-town (2).xlsx' 
sheet_name = 'Colleges'  
column_name = 'Colleges'  

df = pd.read_excel(file_path, sheet_name=sheet_name)

for index, row in df.iterrows():
    institution_name = row[column_name]
    latitude = row['Latitude']
    longitude = row['Longitude']

    if pd.isnull(latitude) or pd.isnull(longitude):
        print(f"Filling geolocation for: {institution_name}")

        location = get_geolocation(api_key, institution_name)

        if location:
            print(f"{institution_name}: Latitude - {location[0]}, Longitude - {location[1]}")
            df.at[index, 'Latitude'] = location[0]
            df.at[index, 'Longitude'] = location[1]
        else:
            print(f"Failed to get geolocation for: {institution_name}")

df.to_excel(file_path, sheet_name=sheet_name, index=False)

