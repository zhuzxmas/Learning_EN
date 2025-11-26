import json
from datetime import datetime
import requests
import configparser
import os
from msal import PublicClientApplication, ConfidentialClientApplication

config = configparser.ConfigParser()
# to check if local file config.cfg is available, for local application
if os.path.exists('./config.cfg'):
    config.read(['config.cfg'])
    azure_settings = config['azure']
    proxy_settings = config['proxy_add']

    client_id = azure_settings['client_id']
    client_secret = azure_settings['client_secret']
    tenant_id = azure_settings['tenant_id']
    username = azure_settings['username']
    team_id = azure_settings['team_id']
    channel_id = azure_settings['channel_id']
    message_id = azure_settings['message_id']
    proxy_add = proxy_settings['proxy_add']
    userId = azure_settings['userId']
    site_id = azure_settings['site_id']
    list_id = azure_settings['list_id']
    item_id = azure_settings['item_id']
    # days_number = int(input("Please enter the number of days to extract the information from Teams Shifts API: \n"))
    deeplx_settings = config['DeepLx']
    key_deeplx = deeplx_settings['secret_key']
else:  # to get this info from Github Secrets, for Github Action running application
    client_id = os.environ['client_id']
    client_secret = os.environ['client_secret']
    tenant_id = os.environ['tenant_id']
    username = os.environ['username']
    userId = os.environ['userId']
    team_id = os.environ['team_id']
    channel_id = os.environ['channel_id']
    message_id = os.environ['message_id']
    site_id = os.environ['site_id']
    list_id = os.environ['list_id']
    item_id = os.environ['item_id']
    openid = os.environ['openid']
    proxy_add = os.environ['proxy_add']
    key_deeplx = os.environ['key_deeplx']

# config.read(['config1.cfg']) # to get the scopes
# azure_settings_scope = config['azure1']
# scope_list = azure_settings_scope['scope_list'].replace(' ','').split(',')

scope_list = ["https://graph.microsoft.com/.default"]
# print( 'Scope List is: ', scope_list, '\n')

proxies = {
    "http": proxy_add,
    "https": proxy_add
}


def get_deeplx_key():
    return key_deeplx


def get_refresh_token_from_SP(access_token, site_id=site_id, list_id=list_id, item_id=item_id):
    # GET /sites/{site-id}/lists/{list-id}/items
    # Replace these with your actual IDs.

    # Construct the URL
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items/{item_id}"

    # Prepare headers
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    # Make the Get request
    try:
        response = requests.get(url, headers=headers)
    except:
        response = requests.get(url, headers=headers, proxies=proxies)

    if response.status_code == 200:
        print("Refresh Token Obtained successfully!")
        Refresh_token = response.json()['fields']['Refresh_Token']
    else:
        Refresh_token = ''
    return Refresh_token


def get_access_token_with_refresh(refresh_token, client_id=client_id, tenant_id=tenant_id):
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

    data = {
        "client_id": client_id,
        "scope": "https://graph.microsoft.com/.default",
        "refresh_token": refresh_token,
        "grant_type": "refresh_token"
    }

    try:
        response = requests.post(url, data=data)
    except:
        response = requests.post(url, data=data, proxies=proxies)

    if response.status_code == 200:
        print("Access Token Obtained successfully!")
        Access_token = response.json()['access_token']
    else:
        Access_token = ''
    return Access_token


def func_login():

    ### to create msal connection ###
    try:
        app = PublicClientApplication(
            client_id=client_id,
            authority='https://login.microsoftonline.com/common',
        )
    except:
        app = PublicClientApplication(
            client_id=client_id,
            authority='https://login.microsoftonline.com/common',
            proxies=proxies
        )

    result = None

    # Firstly, check the cache to see if this end user has signed in before...
    accounts = app.get_accounts(username=username)
    if accounts:
        result = app.acquire_token_silent(scope_list, account=accounts[0])

    if not result:
        print("No suitable token exists in cache. Let's get a new one from Azure AD.")

        flow = app.initiate_device_flow(scopes=scope_list)
        if "user_code" not in flow:
            raise ValueError(
                "Fail to create device flow. Err: %s" % json.dumps(flow, indent=4))

        # print(flow["message"])
        print(
            f"user_code is: {flow['user_code']}, login address: {flow['verification_uri']}")

        # 示例数据
        data = {
            "code": {"value": flow['user_code']},
        }

        message_str1 = flow['user_code']
        message_str2 = flow['verification_uri']

        send_Teams_Channel_Message(message_str2)
        send_Teams_Channel_Message(message_str1)

        # 推送消息
        # result1 = send_template_message(openid, template_id, data)
        # print(result1)  # 打印推送结果

        # Ideally you should wait here, in order to save some unnecessary polling
        # input("Press Enter after signing in from another device to proceed, CTRL+C to abort.")

        result = app.acquire_token_by_device_flow(
            flow)  # By default it will block
        # You can follow this instruction to shorten the block time
        #    https://msal-python.readthedocs.io/en/latest/#msal.PublicClientApplication.acquire_token_by_device_flow
        # or you may even turn off the blocking behavior,
        # and then keep calling acquire_token_by_device_flow(flow) in your own customized loop

        refresh_token = result['refresh_token']
        access_token = result['access_token']

        # Get today's date in ISO format (YYYY-MM-DD or ISO 8601 format)
        # today = datetime.now().isoformat()
        today = datetime.now().strftime('%Y-%m-%d')

        # Please Note: for referesh token, its length is more than 255, so in Microsoft Lists, this column shall be multi-line, not single line
        fields_data = {
                "Refresh_Token": refresh_token,
                "Refresh_Token_Obtained_Date": today,
                "Refresh_Token_Last_Use_Date": today
        }
        
        update_sharepoint_list_item(fields_data, access_token)

    return {'result': result, 'proxies': proxies}


def func_login_secret():
    scopes = ['https://graph.microsoft.com/.default']
    # Create a preferably long-lived app instance which maintains a token cache.
    try:
        app = ConfidentialClientApplication(
            client_id=client_id,
            authority='https://login.microsoftonline.com/{}'.format(tenant_id),
            client_credential=client_secret,
        )
    except:
        app = ConfidentialClientApplication(
            client_id=client_id,
            authority='https://login.microsoftonline.com/{}'.format(tenant_id),
            client_credential=client_secret,
            proxies=proxies
        )
    # Acquire a token using the client credentials flow
    result = None

    # Firstly, checks the cache to see if there is a token it can use
    # If the token is available in the cache, it will return the token
    result = app.acquire_token_silent(scopes=scopes, account=None)

    if not result:
        result = app.acquire_token_for_client(scopes=scopes)

    if "access_token" in result:
        print("Access token got successfully!")
        # print("Access token:", result['access_token'])
    else:
        print(result.get("error"))
        print(result.get("error_description"))
        # You may need this when reporting a bug
        print(result.get("correlation_id"))

    return {'result': result, 'proxies': proxies}


def send_Teams_Channel_Message(message_str, team_id=team_id, channel_id=channel_id, message_id=message_id):

    login_return_app = func_login_secret()
    result_app = login_return_app['result']
    access_token_app = result_app['access_token']
    proxies = login_return_app['proxies']

    refresh_token = get_refresh_token_from_SP(access_token_app)
    access_token = get_access_token_with_refresh(refresh_token)

    # Construct the URL
    # POST /teams/{team-id}/channels/{channel-id}/messages
    url = f"https://graph.microsoft.com/v1.0/teams/{team_id}/channels/{channel_id}/messages/{message_id}/replies"

    # Prepare headers
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    fields_data = {
        "body": {
            "content": message_str
        }
    }

    # Make the Post request
    try:
        response = requests.post(
            url, headers=headers, data=json.dumps(fields_data))
    except:
        response = requests.post(
            url, headers=headers, data=json.dumps(fields_data), proxies=proxies)


def update_sharepoint_list_item(fields_data,access_token, site_id=site_id, list_id=list_id, item_id=item_id):
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

    # Construct the URL
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items/{item_id}/fields"
    # url_columns = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/columns"

    # Prepare headers
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    # # Please Note: for referesh token, its length is more than 255, so in Microsoft Lists, this column shall be multi-line, not single line
    # fields_data = {
    #         "Refresh_Token": refresh_token,
    #         "Refresh_Token_Obtained_Date": today,
    #         "Refresh_Token_Last_Use_Date": today
    # }

    # Make the PATCH request
    try:
        response = requests.patch(
            url, headers=headers, data=json.dumps(fields_data))
    except:
        response = requests.patch(
            url, headers=headers, data=json.dumps(fields_data), proxies=proxies)

    if response.status_code == 200:
        print("Item updated successfully!")
        return response.json()
    else:
        print(f"Error: {response.status_code}")
        print(f"Error message: {response.text}")
        return None
