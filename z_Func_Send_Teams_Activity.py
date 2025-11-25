import json
import requests
import datetime
import os
from pandas import DataFrame
from datetime import datetime
import funcLG
import z_Func_Update_GitHub_Repo_Secrects

# Replace these with your actual IDs.
SITE_ID = "5e9a2fd6-d868-4d52-99e3-2780b185297e"
LIST_ID = "5e116731-b18a-4d7e-ac80-41ebafde4353"
ITEM_ID = "1"

# to login into MS365 and get the return value info.
login_return = funcLG.func_login_secret()
result = login_return['result']
access_token = result['access_token']
proxies = login_return['proxies']

# userTeamwork: sendActivityNotification
# https://learn.microsoft.com/en-us/graph/api/userteamwork-sendactivitynotification?view=graph-rest-1.0&tabs=http

# POST /users/{userId | user-principal-name}/teamwork/sendActivityNotification

def send_Teams_Activity(userId, fields_data):

    # Get access token
    global access_token
    global proxies
    if access_token is None:
        # to login into MS365 and get the return value info.
        login_return = funcLG.func_login_secret()
        result = login_return['result']
        access_token = result['access_token']
        proxies = login_return['proxies']
    else:
        pass

    # Construct the URL
    url = f"https://graph.microsoft.com/v1.0/users/{userId}/teamwork/sendActivityNotification"

    # Prepare headers
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    # Make the PATCH request
    try:
        response = requests.post(
            url, headers=headers, data=json.dumps(fields_data))
    except:
        response = requests.post(
            url, headers=headers, data=json.dumps(fields_data), proxies=proxies)

# Example usage:
if __name__ == "__main__":

    userId = '6dcd9791-b3df-4059-ba8d-1efe369297dc'

    # Data to update
    # Please Note: for referesh token, its length is more than 255, so in Microsoft Lists, this column shall be multi-line, not single line
    fields_data = {
    "topic": {
        "source": "text",
        "value": "aaaa",
        "webUrl": "https://teams.microsoft.com/l"
    },
    "activityType": "taskCreated",
    "previewText": {
        "content": "New Task Created"
    },
}

    # Update the SharePoint list item
    result = send_Teams_Activity(userId, fields_data)

    if result:
        print("Updated item:", result)

