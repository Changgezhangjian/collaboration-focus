import requests
import json
from pprint import pprint

# Set up authentication
auth = ("YOUR_APP_ID", "YOUR_APP_SECRET")
token_url = "https://login.microsoftonline.com/YOUR_TENANT_ID/oauth2/v2.0/token"
token_data = {
    "grant_type": "client_credentials",
    "scope": "https://graph.microsoft.com/.default"
}
token_response = requests.post(token_url, auth=auth, data=token_data)
access_token = token_response.json()["access_token"]
headers = {"Authorization": f"Bearer {access_token}"}

# Create a new Excel workbook
create_workbook_url = "https://graph.microsoft.com/v1.0/me/drive/root/children"
create_workbook_payload = {
    "name": "My Shared Workbook.xlsx",
    "file": {}
}
create_workbook_response = requests.post(create_workbook_url, headers=headers, json=create_workbook_payload)
workbook_id = create_workbook_response.json()["id"]

# Share the workbook with a user
share_url = f"https://graph.microsoft.com/v1.0/drives/{workbook_id}/items/{workbook_id}/invite"
share_payload = {
    "recipients": [
        {
            "email": "user@example.com"
        }
    ],
    "message": "I'm sharing this Excel workbook with you.",
    "requireSignIn": True,
    "sendInvitation": True,
    "roles": ["write"]
}
share_response = requests.post(share_url, headers=headers, json=share_payload)

# Add data to the workbook
range_url = f"https://graph.microsoft.com/v1.0/drives/{workbook_id}/items/{workbook_id}/workbook/worksheets('Sheet1')/range(address='A1:B2')"
range_payload = {
    "values": [
        ["Name", "Age"],
        ["John", 30],
        ["Jane", 25]
    ]
}
range_response = requests.patch(range_url, headers=headers, json=range_payload)

# Print the response
pprint(range_response.json())
