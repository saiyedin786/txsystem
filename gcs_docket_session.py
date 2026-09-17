import requests
from bs4 import BeautifulSoup
from typing import Optional, Dict, Any

class GCSWebSessionManager:
    """
    Python client for BSNL GCS / MantisBT using Session Cookies & HTML Form Submissions.
    Useful when REST API tokens are disabled or not available.
    """

    def __init__(self, base_url: str):
        self.base_url = base_url.rstrip("/")
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        })

    def login(self, username: str, password: str) -> bool:
        """Authenticate and obtain session cookies."""
        login_url = f"{self.base_url}/login.php"
        payload = {
            "username": username,
            "password": password,
            "secure_session": "on"
        }
        res = self.session.post(login_url, data=payload)
        return "login_page.php" not in res.url

    def create_docket(
        self,
        project_id: int,
        category_id: int,
        summary: str,
        description: str,
        severity: int = 50,  # 50 = minor
        priority: int = 30,  # 30 = normal
        handler_id: int = 0,
        custom_fields: Optional[Dict[str, str]] = None
    ) -> bool:
        """Book a new docket via bug_report.php form submit."""
        url = f"{self.base_url}/bug_report.php"
        data = {
            "m_id": 0,
            "project_id": project_id,
            "category_id": category_id,
            "reproducibility": 10,
            "severity": severity,
            "priority": priority,
            "summary": summary,
            "description": description,
            "handler_id": handler_id
        }

        if custom_fields:
            for field_id, val in custom_fields.items():
                data[f"custom_field_{field_id}"] = val

        res = self.session.post(url, data=data)
        return res.status_code == 200

    def resolve_docket(
        self,
        bug_id: int,
        resolution: int = 20,  # 20 = fixed
        bugnote_text: str = ""
    ) -> bool:
        """Resolve a docket via bug_update.php."""
        url = f"{self.base_url}/bug_update.php"
        data = {
            "bug_id": bug_id,
            "status": 80,  # 80 = resolved
            "resolution": resolution,
            "bugnote_text": bugnote_text,
            "action": "resolve"
        }
        res = self.session.post(url, data=data)
        return res.status_code == 200
