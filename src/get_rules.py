import requests
import json

# Replace with your actual token
ACCESS_TOKEN = "YOUR_MICROSOFT_GRAPH_ACCESS_TOKEN"

# --- API Details ---
API_URL = "https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messageRules"
HEADERS = {
    "Authorization": f"Bearer {ACCESS_TOKEN}",
    "Content-Type": "application/json"
}

try:
    response = requests.get(API_URL, headers=HEADERS)
    response.raise_for_status() # Raises an exception for bad status codes (4xx or 5xx)

    rules_data = response.json()
    all_rules = rules_data.get('value', [])

    print(f"Successfully downloaded {len(all_rules)} rules.")

    # Save to a file for backup and external editing
    with open("outlook_rules_backup.json", "w") as f:
        json.dump(all_rules, f, indent=4)
        
    print("Rules saved to outlook_rules_backup.json")

    # Example: Print display name and sequence
    for rule in all_rules:
        print(f"- ID: {rule['id']}, Name: {rule['displayName']}, Sequence: {rule['sequence']}, Enabled: {rule['isEnabled']}")
        
except requests.exceptions.RequestException as e:
    print(f"An error occurred: {e}")