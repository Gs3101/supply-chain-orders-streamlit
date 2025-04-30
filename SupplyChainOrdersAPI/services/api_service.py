import requests
from config import API_BASE_URL

# === קריאת הזמנות מה-API ===
def get_orders():
    try:
        response = requests.get(API_BASE_URL)
        response.raise_for_status()
        return response.json().get("carts", [])
    except Exception as e:
        print(f"❌ API Error (GET): {e}")
        return []

# === יצירת הזמנה חדשה ===
def create_order(data):
    try:
        response = requests.post(API_BASE_URL, json=data)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        print(f"❌ API Error (CREATE): {e}")

# === עדכון הזמנה קיימת ===
def update_order(order_id, data):
    try:
        response = requests.put(f"{API_BASE_URL}/{order_id}", json=data)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        print(f"❌ API Error (UPDATE): {e}")

# === מחיקת הזמנה ===
def delete_order(order_id):
    try:
        response = requests.delete(f"{API_BASE_URL}/{order_id}")
        response.raise_for_status()
        return response.status_code == 200
    except Exception as e:
        print(f"❌ API Error (DELETE): {e}")

# === בדיקת זמינות API ===
def check_api_health():
    try:
        response = requests.get(API_BASE_URL, timeout=5)
        if response.status_code == 200:
            print("✅ API is reachable.")
        else:
            print(f"⚠️ API returned status code {response.status_code}")
    except Exception as e:
        print(f"❌ API Health Check Failed: {e}")
