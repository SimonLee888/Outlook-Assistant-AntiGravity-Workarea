import base64
import requests


STORE_QUEUE = "***"
API_KEY = "***"
BASE_URL = "https://api.rfms.online/v2"


def get_session_token():
    auth_string = f"{STORE_QUEUE}:{API_KEY}"
    auth_base64 = base64.b64encode(auth_string.encode()).decode()
    headers = {"Authorization": f"Basic {auth_base64}"}

    response = requests.post(f"{BASE_URL}/session/begin", headers=headers, timeout=30)
    response.raise_for_status()
    return response.json().get("sessionToken")


def get_estimate_data(estimate_number, session_token):
    headers = {
        "Authorization": (
            "Basic "
            + base64.b64encode(f"{STORE_QUEUE}:{session_token}".encode()).decode()
        ),
        "Content-Type": "application/json",
    }

    response = requests.get(
        f"{BASE_URL}/estimate/{estimate_number}?locked=false",
        headers=headers,
        timeout=30,
    )
    response.raise_for_status()
    result = response.json().get("result") or {}

    sold_to = result.get("soldTo", {})
    lines = result.get("lines", [])
    total_area = sum(
        line.get("quantity", 0)
        for line in lines
        if line.get("saleUnits", "").upper() == "SF"
    )

    return {
        "estimate_number": estimate_number,
        "job_number": result.get("jobNumber"),
        "customer_name": sold_to.get("businessName") or sold_to.get("lastName"),
        "number_of_lines": len(lines),
        "total_area": total_area,
        "created_date": (result.get("quoteDate") or result.get("orderDate") or "").split("T")[0]
        or None,
    }
