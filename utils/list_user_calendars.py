import os
import json
import requests
import msal



# Azure client/app ID
azure_client_id = "client-id"
# Azure Tenant ID
azure_tenant_id = "tenant-id"
# Azure user's ID
user_id = "user-id"

ms_authority = f"https://login.microsoftonline.com/{azure_tenant_id}"
ms_scopes = [
        "https://graph.microsoft.com/Calendars.ReadWrite", 
        "https://graph.microsoft.com/Calendars.ReadWrite.Shared", 
        "https://graph.microsoft.com/Group.ReadWrite.All", 
        "https://graph.microsoft.com/offline_access"
    ]

cache_file = "token_cache.json"



def get_access_token() -> str:
    """ Retrieves access token from cache or authenticates user if needed """
    
    # Load token cache if exists
    cache = msal.SerializableTokenCache()
    if os.path.exists(cache_file):
        with open(cache_file, "r") as fp:
            cache.deserialize(fp.read())

    graph_app = msal.PublicClientApplication(azure_client_id, authority=ms_authority, token_cache=cache)

    # Try to get token silently (without asking user)
    accounts = graph_app.get_accounts()
    if accounts:
        result = graph_app.acquire_token_silent(ms_scopes, account=accounts[0])

    else:
        # If no valid token, ask the user to log in interactively
        result = graph_app.acquire_token_interactive(ms_scopes)

    # Save updated token cache
    with open(cache_file, "w") as fp:
        fp.write(cache.serialize())

    if "access_token" in result:
        return result["access_token"]

    else:
        err = "Failed to authenticate: " + json.dumps(result, indent=4)
        raise Exception(err)


access_token = get_access_token()

# Microsoft Graph API endpoint
url = f"https://graph.microsoft.com/v1.0/users/{user_id}/calendars"

# Headers
headers = {
    "Authorization": f"Bearer {access_token}",
    "Content-Type": "application/json"
}

# Make the request
response = requests.get(url, headers=headers)

# Print the calendar list
if response.status_code == 200:
    calendars = response.json()
    for calendar in calendars.get("value", []):
        print(f"Calendar Name: {calendar['name']}, ID: {calendar['id']}")
else:
    print(f"Error: {response.status_code}, {response.text}")

