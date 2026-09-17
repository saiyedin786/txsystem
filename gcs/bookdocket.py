import os
import json
import base64
import requests
from typing import Optional

# 1. GCS Server Configuration & API Token
GCS_BASE_URL = "https://intranetguj.bsnl.co.in/gcs"
API_TOKEN = "api_token"  # Replace with your actual token

# 2. Optional Excel File Attachment Path
EXCEL_FILE_PATH: Optional[str] = None  # e.g., "C:/path/to/my_data.xlsx" or None

# 3. Define Docket Payload
new_docket_payload = {
    "summary": "CREATE TCS 4G SERVICE FOR Rajkot to Amreli",
    "description": "Request for creating 4G TCS service link connectivity between Rajkot to Amreli",
    "project": {
        "name": "GUJ_TX_NOC"
    },
    "category": {
        "name": "complaint"
    },
    "priority": {
        "name": "minor"
    },
    "severity": {
        "name": "minor"
    },
    "handler": {
        "name": "gjtxnoc"  # Assigned handler username
    },
    "custom_fields": [
        {
            "field": {
                "name": "SSA_NAME"
            },
            "value": "Rajkot"
        },
        {
            "field": {
                "name": "Amount_of_work"
            },
            "value": "1"
        }
    ]
}

# 4. Attach Excel File (If provided and file exists)
if EXCEL_FILE_PATH and os.path.isfile(EXCEL_FILE_PATH):
    with open(EXCEL_FILE_PATH, "rb") as f:
        encoded_content = base64.b64encode(f.read()).decode("utf-8")
    
    new_docket_payload["files"] = [
        {
            "name": os.path.basename(EXCEL_FILE_PATH),
            "content": encoded_content
        }
    ]
    print(f"📎 Attaching file: {os.path.basename(EXCEL_FILE_PATH)}")

# 5. Headers
headers = {
    "Authorization": API_TOKEN,
    "Content-Type": "application/json"
}

# 6. POST Request to Create Docket
url = f"{GCS_BASE_URL}/api/rest/issues/"
response = requests.post(url, headers=headers, json=new_docket_payload)

# 7. Check Output
if response.status_code == 201:  # HTTP 201 Created
    created_issue = response.json().get("issue", {})
    issue_id = created_issue.get("id")
    print(" SUCCESS! Docket Booked.")
    print(f"New Docket ID:   {issue_id}")
    print(f"Summary:         {created_issue.get('summary')}")
    print(f"Status:          {created_issue.get('status', {}).get('name')}")
else:
    print(f" Error Booking Docket! Status Code: {response.status_code}")
    print(response.text)