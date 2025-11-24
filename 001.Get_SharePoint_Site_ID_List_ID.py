import json, requests, datetime, os
from pandas import DataFrame
from datetime import datetime
import funcLG

# login_return = funcLG.func_login() # to login into MS365 and get the return value info.
# result = login_return['result']
# refresh_token = result['refresh_token']
# proxies = login_return['proxies']

login_return_secret = funcLG.func_login_secret() # to login into MS365 and get the return value info.
result_secret = login_return_secret['result']
proxies = login_return_secret['proxies']

# https://learn.microsoft.com/en-us/graph/api/site-getallsites?view=graph-rest-1.0&tabs=http

# Get access token
access_token = result_secret['access_token']

# Construct the URL
url = f"https://graph.microsoft.com/v1.0/sites/getAllSites"

# Prepare headers
headers = {
    'Authorization': f'Bearer {access_token}',
    'Content-Type': 'application/json'
}

try:
    # Make the Get request
    response = requests.get(url, headers=headers)
except:
    response = requests.get(url, headers=headers, proxies=proxies)
    
if response.status_code == 200:
    print("Item updated successfully!")
else:
    print(f"Error: {response.status_code}")
    print(f"Error message: {response.text}")

# Example usage:
if __name__ == "__main__":

    # Get today's date in ISO format (YYYY-MM-DD or ISO 8601 format)
    today = datetime.now().strftime('%Y-%m-%d') 
