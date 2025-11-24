import json
import requests
import datetime
import os
from pandas import DataFrame
from datetime import datetime
import funcLG

# Replace these with your actual IDs.
SITE_ID = "5e9a2fd6-d868-4d52-99e3-2780b185297e"
LIST_ID = "5e116731-b18a-4d7e-ac80-41ebafde4353"
ITEM_ID = "1"

# to login into MS365 and get the return value info.
login_return = funcLG.func_login()
result = login_return['result']
refresh_token = result['refresh_token']
access_token = result['access_token']
proxies = login_return['proxies']

# # to login into MS365 and get the return value info.
# login_return_secret = funcLG.func_login_secret()
# result_secret = login_return_secret['result']

# PATCH /sites/{site-id}/lists/{list-id}/items/{item-id}/fields
# https://learn.microsoft.com/en-us/graph/api/listitem-update?view=graph-rest-1.0&tabs=http


def update_sharepoint_list_item(site_id, list_id, item_id, fields_data):
    """
    Update a SharePoint list item using Microsoft Graph API

    Args:
        site_id (str): The SharePoint site ID
        list_id (str): The SharePoint list ID
        item_id (str): The SharePoint list item ID
        fields_data (dict): Dictionary containing the fields to update

    Returns:
        dict: Response from the API
    """
    # Get access token
    if access_token is None:
        # to login into MS365 and get the return value info.
        login_return = funcLG.func_login()
        result = login_return['result']
        access_token = result['access_token']
        proxies = login_return['proxies']
    else:
        pass


    # Construct the URL
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items/{item_id}/fields"
    # url_columns = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/columns"

    # Prepare headers
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    # Prepare the payload
    payload = fields_data

    # Make the PATCH request
    try:
        response = requests.patch(
            url, headers=headers, data=json.dumps(payload))
    except:
        response = requests.patch(
            url, headers=headers, data=json.dumps(payload), proxies=proxies)

    if response.status_code == 200:
        print("Item updated successfully!")
        return response.json()
    else:
        print(f"Error: {response.status_code}")
        print(f"Error message: {response.text}")
        return None


# Example usage:
if __name__ == "__main__":

    # Get today's date in ISO format (YYYY-MM-DD or ISO 8601 format)
    today = datetime.now().isoformat()

    # Data to update
    fields_to_update = {
        "fields": {
            "Refresh_Token": refresh_token,
            "Refresh_Token_Last_Use_Date": today
        }}

    # Update the SharePoint list item
    result = update_sharepoint_list_item(
        SITE_ID, LIST_ID, ITEM_ID, fields_to_update)

    if result:
        print("Updated item:", result)
