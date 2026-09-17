import json
import requests

# 1. Server Configuration
GCS_BASE_URL = "https://intranetguj.bsnl.co.in/gcs"
API_TOKEN = "api_token"  # Replace with your actual token
DOCKET_ID = 86457  # Replace with the Docket ID you want to view

# 2. Setup Request Headers
headers = {
    "Authorization": API_TOKEN,
    "Content-Type": "application/json"
}

# 3. Call REST API to View Docket
url = f"{GCS_BASE_URL}/api/rest/issues/{DOCKET_ID}"
response = requests.get(url, headers=headers)

# 4. Process and Print Response
if response.status_code == 200:
    docket_data = response.json()
    issue = docket_data["issues"][0]
    
    print("=== DOCKET DETAILS ===")
    print(f"ID:          {issue['id']}")
    print(f"Summary:     {issue['summary']}")
    print(f"Status:      {issue['status']['name']}")
    print(f"Severity:    {issue['severity']['name']}")
    print(f"Priority:    {issue['priority']['name']}")
    print(f"Reporter:    {issue['reporter']['name']}")
    print(f"Project:     {issue['project']['name']}")
    print(f"Date Created:{issue['created_at']}")
    print(f"Description: {issue['description']}")
    
    # Custom Fields (e.g. SSA_NAME)
    if "custom_fields" in issue:
        print("\nCustom Fields:")
        for cf in issue["custom_fields"]:
            print(f" - {cf['field']['name']}: {cf['value']}")
else:
    print(f"Failed to fetch docket. Status Code: {response.status_code}")
    print(response.text)