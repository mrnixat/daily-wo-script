import requests
from msal import ConfidentialClientApplication
import os
import base64
import time 
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import settings.config as cfg
# Graph API endpoints
AUTHORITY = f'https://login.microsoftonline.com/{cfg.ten_id}'
SCOPE = ['https://graph.microsoft.com/.default']
GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0'

# === AUTHENTICATE ===
app = ConfidentialClientApplication(
    client_id=cfg.cli_id,
    client_credential=cfg.secret_pas,
    authority=AUTHORITY
)

token_response = app.acquire_token_for_client(scopes=SCOPE)

if 'access_token' not in token_response:
    print("Authentication failed:")
    print(token_response.get("error_description"))
    exit()

access_token = token_response['access_token']
headers = {
    'Authorization': f'Bearer {access_token}',
    'Accept': 'application/json'
}


SAVE_DIR = os.path.expanduser("~/Downloads/Hapag_XTs")
os.makedirs(SAVE_DIR, exist_ok=True)



def get_folder_id(folder_name):
    folders_url = f"{GRAPH_ENDPOINT}/users/{cfg.sh_mailbox}/mailFolders?$top=100"
    folders_resp = requests.get(folders_url, headers=headers)
    if folders_resp.status_code == 200:
        for folder in folders_resp.json().get('value', []):
            if folder.get('displayName', '').lower() == folder_name.lower():
                return folder['id']
    print(f"Could not find '{folder_name}' folder.")
    return None

# --- Only fetch messages from Actioned received today ---

def fetch_and_save_wosd_pdfs(folder_name, save_folder=None, start_str=None, end_str=None, start_dt=None, end_dt=None):
    folder_id = get_folder_id(folder_name)
    if not folder_id:
        return
    central = ZoneInfo("America/Chicago")
    if not start_str or not end_str:
        # Use start_dt/end_dt if provided, else default to today in Central
        if start_dt is None or end_dt is None:
            now = datetime.now(central)
            start_dt = now.replace(hour=0, minute=0, second=0, microsecond=0)
            end_dt = now.replace(hour=23, minute=59, second=59, microsecond=999999)
        # Convert to UTC for Graph API
        start_str = start_dt.astimezone(ZoneInfo("UTC")).strftime('%Y-%m-%dT%H:%M:%SZ')
        end_str = end_dt.astimezone(ZoneInfo("UTC")).strftime('%Y-%m-%dT%H:%M:%SZ')
    if not save_folder:
        save_folder = os.path.expanduser("~/Downloads/Hapag_XTs")

    messages_url = (
        f"{GRAPH_ENDPOINT}/users/{cfg.sh_mailbox}/mailFolders/{folder_id}/messages"
        f"?$filter=receivedDateTime ge {start_str} and receivedDateTime le {end_str}"
        f"&$top=100&$orderby=receivedDateTime desc"
    )
    count = 0
    msg_seen = 0
    while messages_url:
        messages_resp = requests.get(messages_url, headers=headers)
        if messages_resp.status_code != 200:
            print(f"Failed to fetch messages from {folder_name}:", messages_resp.status_code, messages_resp.text)
            break
        data = messages_resp.json()
        messages = data.get('value', [])
        for msg in messages:
            subject = msg.get('subject', '') or ''
            if 'wosd' not in subject.lower():
                print("Skipping email with subject:", subject)
                continue
            msg_seen += 1
            msg_id = msg['id']
            att_url = f"{GRAPH_ENDPOINT}/users/{cfg.sh_mailbox}/messages/{msg_id}/attachments"
            att_resp = requests.get(att_url, headers=headers)
            if att_resp.status_code != 200:
                print(f"Failed to fetch attachments for message {msg_id}")
                continue
            found_any = False
            for att in att_resp.json().get('value', []):
                name = att.get('name', '')
                if name.lower().endswith('.pdf') and name.startswith('WOSD'):
                    found_any = True
                    file_path = os.path.join(save_folder, name)
                    if os.path.exists(file_path):
                        print(f"Skipping duplicate: {name}")
                        continue
                    content_bytes = att.get('contentBytes')
                    if not content_bytes:
                        print(f"Attachment {name} has no contentBytes (may be too large or not a file attachment)")
                        continue
                    try:
                        with open(file_path, 'wb') as f:
                            f.write(base64.b64decode(content_bytes))
                        print(f"Saved: {name}")
                        count += 1
                    except Exception as e:
                        print(f"Failed to save {name}: {e}")
            if not found_any:
                msg_time = msg.get('receivedDateTime', 'unknown time')
                print(f"No WOSD PDF attachments found in message sent at {msg_time}")
        messages_url = data.get('@odata.nextLink')
    print(f"Checked {msg_seen} messages. Downloaded {count} WOSD PDFs from {folder_name}.")


# Search Inbox first, then Actioned (today, Central time)
"""
central = ZoneInfo("America/Chicago")
now = datetime.now(central)
start_today = now.replace(hour=0, minute=0, second=0, microsecond=0)
end_today = now.replace(hour=23, minute=59, second=59, microsecond=999999)

print("Now grabbing emails from Inbox")
print("")
time.sleep(2)
fetch_and_save_wosd_pdfs('Inbox', start_dt=start_today, end_dt=end_today)

print("Now grabbing emails from Actioned")
print("")
time.sleep(2)
fetch_and_save_wosd_pdfs('Actioned', start_dt=start_today, end_dt=end_today)

# SEARCHING YESTERDAY'S EMAILS (Central time, 4PM-11:59PM)
prev_day = now - timedelta(days=1)
start_prev = prev_day.replace(hour=16, minute=0, second=0, microsecond=0)
end_prev = prev_day.replace(hour=23, minute=59, second=59, microsecond=999999)
save_folder = os.path.expanduser("~/Downloads/Hapag_XTs/Hapag_XTs_Yesterday")
os.makedirs(save_folder, exist_ok=True)

print("Now grabbing emails from Actioned for yesterday from 4:00 PM to 11:59 PM")
print("")
time.sleep(2)
fetch_and_save_wosd_pdfs('Actioned', save_folder=save_folder, start_dt=start_prev, end_dt=end_prev)
"""