# Code explained

> **Source**: Google Gemini

- [Programmatic Rule Management](https://gemini.google.com/share/814138c64099)

Yes, you can manage Outlook rules programmatically using the
**Microsoft Graph API**. This is the best method for the bulk download,
update, and re-upload process you described, as the web interface doesn't
offer direct export.

Since Python is requested and is excellent for scripting API calls, here are
examples using Python.

The most common way to interact with the Microsoft Graph API in Python is
using the `requests` library to send HTTP requests.

The endpoint you'll use is:

`https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messageRules`

This process requires a Microsoft Entra ID (formerly Azure AD) application
registration to get an **Access Token** with the `MailboxSettings.ReadWrite`
permission. The example below assumes you've already authenticated and have
a valid access token.

## 1\. Download (Get) All Rules

This script retrieves all your inbox rules and prints their key details.

```python
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

```

## 2\. Update a Single Rule (Example: Change Sequence Order)

To update a rule, you use the `PATCH` method with the rule's ID and only send the properties you want to change.

In this example, we assume you've identified a rule's ID and want to change its execution order (`sequence`).

```python
import requests

# Use the same token and headers from the previous example
ACCESS_TOKEN = "YOUR_MICROSOFT_GRAPH_ACCESS_TOKEN"
HEADERS = {
    "Authorization": f"Bearer {ACCESS_TOKEN}",
    "Content-Type": "application/json"
}

# Replace with the actual ID of the rule you want to modify
RULE_ID_TO_UPDATE = "AAMkADk0NGE5YjIxLWM1ZTItNDc3OS04YmRhLTc2YmFkMjM3ZjcxMABHAAAA"
NEW_SEQUENCE_ORDER = 1

# --- API Details ---
API_URL_UPDATE = f"https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messageRules/{RULE_ID_TO_UPDATE}"

# Payload only contains the property we want to change
update_payload = {
    "sequence": NEW_SEQUENCE_ORDER
}

try:
    response = requests.patch(API_URL_UPDATE, headers=HEADERS, json=update_payload)
    response.raise_for_status()

    print(f"Rule ID {RULE_ID_TO_UPDATE} successfully updated.")
    print(f"New sequence order is: {response.json().get('sequence')}")

except requests.exceptions.RequestException as e:
    print(f"An error occurred during update: {e}")
```

## 3\. Create a New Rule (Example: Move Mail from a Specific Sender)

To create a new rule, you use the `POST` method. The payload must include the rule's `displayName`, `conditions`, and `actions`.

```python
import requests

# Use the same token and headers from the previous example
ACCESS_TOKEN = "YOUR_MICROSOFT_GRAPH_ACCESS_TOKEN"
HEADERS = {
    "Authorization": f"Bearer {ACCESS_TOKEN}",
    "Content-Type": "application/json"
}

# Replace with the email address and target folder ID
SENDER_EMAIL = "specific.sender@example.com"
TARGET_FOLDER_ID = "AAMkADk0NGE5YjIxLWM1ZTItNDc3OS04YmRhLTc2YmFkMjM3ZjcxMAAuAAAA..." # ID of the folder where mail should be moved

# --- API Details ---
API_URL_CREATE = "https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messageRules"

# Rule payload
new_rule_payload = {
    "displayName": f"Move Mail from {SENDER_EMAIL}",
    "sequence": 5,
    "isEnabled": True,
    "conditions": {
        "senderContains": [
            SENDER_EMAIL
        ]
    },
    "actions": {
        "moveToFolder": TARGET_FOLDER_ID
    }
}

try:
    response = requests.post(API_URL_CREATE, headers=HEADERS, json=new_rule_payload)
    response.raise_for_status()

    new_rule = response.json()
    print(f"New rule created successfully: {new_rule.get('displayName')} (ID: {new_rule.get('id')})")

except requests.exceptions.RequestException as e:
    print(f"An error occurred during rule creation: {e}")
```

-----

## How to Perform Bulk Updates

To achieve your goal of downloading, updating some, and re-uploading:

1. **Run Example 1** to get all rules and save them to a `.json` file.
1. **Manually Edit the JSON:** Open the `outlook_rules_backup.json` file.
Identify the rules you want to change, and modify their properties
(e.g., change the `sequence` value).
1. **Run a Script to Iterate and Update:** You would write a single Python
script that:

    - Loads the **edited** rules file (`outlook_rules_backup.json`).
    - **Deletes all existing rules** using the Graph API (`DELETE` on each rule ID).
    - **Iterates** through your edited list of rules and **re-creates them all** using the Graph API (`POST`).

This "delete all, then re-create all" method is often the simplest and safest way to perform a bulk update, as it ensures the server-side state matches your local, edited state.
