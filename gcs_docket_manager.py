import os
import json
import logging
import requests
from typing import Dict, Any, Optional

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger("GCSDocketManager")

class GCSDocketManager:
    """
    Python client for BSNL Gujarat Collaborative System (GCS / MantisBT)
    to Create (Book), View, Update, and Resolve Dockets via MantisBT REST API.
    """

    def __init__(self, base_url: str, api_token: str):
        """
        Initialize the GCS Docket Manager.

        :param base_url: Base URL of GCS installation (e.g. 'https://intranetguj.bsnl.co.in/gcs')
        :param api_token: User API token generated from MantisBT My Account -> API Tokens
        """
        self.base_url = base_url.rstrip("/")
        self.api_token = api_token
        self.headers = {
            "Authorization": self.api_token,
            "Content-Type": "application/json"
        }

    def create_docket(
        self,
        summary: str,
        description: str,
        project_name: str = "GUJ_TX_NOC",
        category_name: str = "Service",
        priority: str = "normal",
        severity: str = "minor",
        handler_name: Optional[str] = "gjtxnoc",
        custom_fields: Optional[Dict[str, str]] = None,
        file_path: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Book / Create a new docket in GCS.

        :param summary: Brief summary of the docket
        :param description: Detailed description of the docket issue
        :param project_name: Project name in GCS (e.g. 'GUJ_TX_NOC')
        :param category_name: Category name (e.g. 'Service')
        :param priority: Priority level ('low', 'normal', 'high', 'urgent', 'immediate')
        :param severity: Severity level ('feature', 'trivial', 'text', 'tweak', 'minor', 'major', 'crash', 'block')
        :param handler_name: Username of the assigned handler/person (e.g. 'gjbxnoc')
        :param custom_fields: Dictionary of custom fields, e.g. {"SSA_NAME": "Junagadh", "Amount_of_work": "1"}
        :param file_path: Optional path to an Excel sheet or file to attach (e.g. 'C:/path/data.xlsx')
        :return: JSON response containing created issue details including issue ID
        """
        url = f"{self.base_url}/api/rest/issues/"

        payload: Dict[str, Any] = {
            "summary": summary,
            "description": description,
            "project": {"name": project_name},
            "category": {"name": category_name},
            "priority": {"name": priority},
            "severity": {"name": severity}
        }

        if handler_name:
            payload["handler"] = {"name": handler_name}

        if custom_fields:
            payload["custom_fields"] = [
                {"field": {"name": k}, "value": v} for k, v in custom_fields.items()
            ]

        if file_path and os.path.isfile(file_path):
            import base64
            with open(file_path, "rb") as f:
                encoded = base64.b64encode(f.read()).decode("utf-8")
            payload["files"] = [
                {
                    "name": os.path.basename(file_path),
                    "content": encoded
                }
            ]
            logger.info(f"Attaching file '{os.path.basename(file_path)}' to docket creation payload.")

        try:
            response = requests.post(url, headers=self.headers, json=payload)
            response.raise_for_status()
            data = response.json()
            issue_id = data.get("issue", {}).get("id")
            logger.info(f"Docket successfully created! Docket ID: {issue_id}")
            return data
        except requests.exceptions.RequestException as e:
            logger.error(f"Failed to create docket: {e}")
            if e.response is not None:
                logger.error(f"Response status: {e.response.status_code}, body: {e.response.text}")
            raise

    def view_docket(self, docket_id: int) -> Dict[str, Any]:
        """
        View details of a specific docket by ID.

        :param docket_id: The numeric ID of the docket (e.g. 86457)
        :return: JSON object containing docket details
        """
        url = f"{self.base_url}/api/rest/issues/{docket_id}"

        try:
            response = requests.get(url, headers=self.headers)
            response.raise_for_status()
            data = response.json()
            logger.info(f"Successfully fetched details for Docket #{docket_id}")
            return data
        except requests.exceptions.RequestException as e:
            logger.error(f"Failed to fetch docket #{docket_id}: {e}")
            if e.response is not None:
                logger.error(f"Response status: {e.response.status_code}, body: {e.response.text}")
            raise

    def update_docket(
        self,
        docket_id: int,
        summary: Optional[str] = None,
        description: Optional[str] = None,
        priority: Optional[str] = None,
        severity: Optional[str] = None,
        handler_name: Optional[str] = None,
        custom_fields: Optional[Dict[str, str]] = None
    ) -> Dict[str, Any]:
        """
        Update attributes of an existing docket.

        :param docket_id: Numeric ID of the docket to update
        :param summary: Updated summary text (optional)
        :param description: Updated description (optional)
        :param priority: Updated priority (optional)
        :param severity: Updated severity (optional)
        :param handler_name: Assign to new handler username (optional)
        :param custom_fields: Dictionary of updated custom fields (optional)
        :return: JSON response from GCS server
        """
        url = f"{self.base_url}/api/rest/issues/{docket_id}"

        payload: Dict[str, Any] = {}

        if summary is not None:
            payload["summary"] = summary
        if description is not None:
            payload["description"] = description
        if priority is not None:
            payload["priority"] = {"name": priority}
        if severity is not None:
            payload["severity"] = {"name": severity}
        if handler_name is not None:
            payload["handler"] = {"name": handler_name}
        if custom_fields:
            payload["custom_fields"] = [
                {"field": {"name": k}, "value": v} for k, v in custom_fields.items()
            ]

        if not payload:
            logger.warning("No update fields provided.")
            return {}

        try:
            response = requests.patch(url, headers=self.headers, json=payload)
            response.raise_for_status()
            data = response.json()
            logger.info(f"Successfully updated Docket #{docket_id}")
            return data
        except requests.exceptions.RequestException as e:
            logger.error(f"Failed to update docket #{docket_id}: {e}")
            if e.response is not None:
                logger.error(f"Response status: {e.response.status_code}, body: {e.response.text}")
            raise

    def resolve_docket(
        self,
        docket_id: int,
        resolution: str = "fixed",
        note_text: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Resolve a docket in GCS.

        :param docket_id: Numeric ID of the docket to resolve
        :param resolution: Resolution code ('fixed', 'reopened', 'unable to reproduce', 'not fixable', 'duplicate', 'no change required', 'suspended', 'wont fix')
        :param note_text: Resolution note or closing remarks (optional)
        :return: JSON response
        """
        url = f"{self.base_url}/api/rest/issues/{docket_id}"

        payload = {
            "status": {"name": "resolved"},
            "resolution": {"name": resolution}
        }

        try:
            # 1. Change status to resolved
            response = requests.patch(url, headers=self.headers, json=payload)
            response.raise_for_status()
            result = response.json()

            # 2. Add resolution note if provided
            if note_text:
                self.add_note(docket_id, note_text)

            logger.info(f"Docket #{docket_id} resolved successfully with status '{resolution}'.")
            return result
        except requests.exceptions.RequestException as e:
            logger.error(f"Failed to resolve docket #{docket_id}: {e}")
            if e.response is not None:
                logger.error(f"Response status: {e.response.status_code}, body: {e.response.text}")
            raise

    def add_note(self, docket_id: int, note_text: str) -> Dict[str, Any]:
        """
        Add a comment / note to a docket.

        :param docket_id: Numeric ID of the docket
        :param note_text: Comment text to add
        :return: JSON response
        """
        url = f"{self.base_url}/api/rest/issues/{docket_id}/notes"

        payload = {
            "text": note_text,
            "view_state": {"name": "public"}
        }

        try:
            response = requests.post(url, headers=self.headers, json=payload)
            response.raise_for_status()
            data = response.json()
            logger.info(f"Added note to Docket #{docket_id}")
            return data
        except requests.exceptions.RequestException as e:
            logger.error(f"Failed to add note to docket #{docket_id}: {e}")
            raise


# =====================================================================
# Demonstration & Usage Example
# =====================================================================
if __name__ == "__main__":
    # Update base URL and API token according to your GCS configuration
    GCS_BASE_URL = os.getenv("GCS_BASE_URL", "https://intranetguj.bsnl.co.in/gcs")
    GCS_API_TOKEN = os.getenv("GCS_API_TOKEN", "YOUR_API_TOKEN_HERE")

    client = GCSDocketManager(base_url=GCS_BASE_URL, api_token=GCS_API_TOKEN)

    print("=== GCS Docket Manager API Example ===")
    print("Replace GCS_BASE_URL and GCS_API_TOKEN with your actual credentials.\n")

    # Example 1: Create (Book) Docket
    # new_docket = client.create_docket(
    #     summary="CREATE TCS 4G SERVICE FOR DIU & WANAKBARA",
    #     description="4G Node integration and service setup for Diu region.",
    #     project_name="GUJ_TX_NOC",
    #     category_name="Service",
    #     priority="normal",
    #     severity="minor",
    #     custom_fields={"SSA_NAME": "Junagadh"}
    # )

    # Example 2: View Docket
    # docket_info = client.view_docket(docket_id=86457)
    # print(json.dumps(docket_info, indent=2))

    # Example 3: Update Docket
    # client.update_docket(
    #     docket_id=86457,
    #     priority="high",
    #     handler_name="gjbxnoc",
    #     custom_fields={"SSA_NAME": "Junagadh"}
    # )

    # Example 4: Resolve Docket
    # client.resolve_docket(
    #     docket_id=86457,
    #     resolution="fixed",
    #     note_text="4G Service configured and tested successfully."
    # )
