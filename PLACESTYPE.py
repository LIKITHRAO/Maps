import pandas as pd
import requests
import openpyxl
import time
import re

def get_place_details(api_key, place_id):
    base_url = "https://maps.googleapis.com/maps/api/place/details/json"
    params = {
        "place_id": place_id,
        "fields": "formatted_phone_number,website",
        "key": api_key
    }

    response = requests.get(base_url, params=params)
    if response.status_code == 200:
        return response.json().get("result", {})
    else:
        print(f"Error: Unable to get details for place ID {place_id}, Status Code: {response.status_code}")
        return {}

def get_places_data(api_key, pincode):
    base_url = "https://maps.googleapis.com/maps/api/place/textsearch/json"
    query = f"Engineering Colleges in {pincode}"
    params = {
        "query": query,
        "key": api_key
    }

    results = []
    while True:
        response = requests.get(base_url, params=params)
        if response.status_code == 200:
            data = response.json()
            for place in data.get("results", []):
                # Filter results to only include those with the exact pincode in the address
                if re.search(rf"\b{pincode}\b", place.get("formatted_address", "")):
                    results.append(place)
            # Handle pagination
            next_page_token = data.get("next_page_token")
            if next_page_token:
                # Google recommends waiting a few seconds before using the next_page_token
                time.sleep(2)
                params['pagetoken'] = next_page_token
            else:
                break
        else:
            print(f"Error: Unable to get data for pincode {pincode}, Status Code: {response.status_code}")
            break

    return results

def extract_data(api_key, place, pincode):
    address = place.get("formatted_address", "")
    
    # Extract state assuming it's the second last part of the address
    address_parts = address.split(", ")
    state = address_parts[-2] if len(address_parts) > 1 else ""

    # Get place details for phone number and website
    place_details = get_place_details(api_key, place.get("place_id"))

    # Convert ratings to stars (assuming ratings are out of 5)
    rating = place.get("rating", 0)
    stars = f"{rating} ⭐" if rating else "No rating available"

    # Determine the school type (college or school)
    school_type = ""
    types = place.get("types", [])
    if any(term in types for term in ["university", "college"]):
        school_type = "College"
    elif any(term in types for term in ["school", "academy"]):
        school_type = "School"

    return {
        "Pincode": pincode,
        "School Name": place.get("name"),
        "Address": address,
        "Reviews": stars,
        "State": state,
        "Phone No": place_details.get("formatted_phone_number", "N/A"),
        "Email ID": "N/A",  # Not provided by Places API
        "School Type": school_type,
        "Website": place_details.get("website", "N/A")
    }


# Modify main function to pass the api_key to extract_data
def main(input_file, output_file, api_key):
    # Load pincodes from input Excel sheet
    df_input = pd.read_excel(input_file)
    
    # Ensure the 'Pincode' column exists in the input file
    if 'Pincode' not in df_input.columns:
        raise KeyError("The input file must contain a column named 'Pincode'.")
    
    # Create an empty DataFrame to store the output
    columns = [
        "Pincode", "School Name", "Address", "Reviews", "State", "Phone No", "Email ID", "School Type", "Website"
    ]
    df_output = pd.DataFrame(columns=columns)
    
    # Process each pincode and fetch data
    data_list = []  # Use a list to collect data before creating DataFrame
    for index, row in df_input.iterrows():
        pincode = row['Pincode']
        print(f"Fetching data for pincode {pincode}...")
        places = get_places_data(api_key, pincode)
        for place in places:
            data_list.append(extract_data(api_key, place, pincode))
    
    # Create DataFrame from the collected data
    df_output = pd.DataFrame(data_list, columns=columns)
    
    # Save the data to output Excel sheet
    df_output.to_excel(output_file, index=False)
    print(f"Data has been successfully saved to {output_file}")


if __name__ == "__main__":
    # Replace with your Google API key
    api_key ="AIzaSyCxrgn6ZZL3IsY_3xrSqQJi_3yT_OKr-n0"
    
    # Input and output file paths
    input_file = "C:\\Users\\likit\\Downloads\\abdn.xlsx"
    output_file = "C:\\Users\\likit\\Downloads\\abdnans.xlsx"
    
    # Run the main function
    main(input_file, output_file, api_key)
