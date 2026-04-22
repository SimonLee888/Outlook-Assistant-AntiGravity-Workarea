import requests


TENANT_ID = "***"
CLIENT_ID = "***"
CLIENT_SECRET = "***"
TEAM_ID = "***"
GRAPH_URL = "https://graph.microsoft.com/beta"


def get_access_token():
    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
    }
    response = requests.post(token_url, data=data, timeout=30)
    response.raise_for_status()
    return response.json()["access_token"]


def get_time_cards(access_token):
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }
    response = requests.get(
        f"{GRAPH_URL}/teams/{TEAM_ID}/schedule/timeCards",
        headers=headers,
        timeout=30,
    )
    response.raise_for_status()
    return response.json().get("value", [])
