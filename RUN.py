from grab_xts import fetch_and_save_wosd_pdfs
from clean_xts import clean_xts
from send_email import send_email
import time
import os
import shutil
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo  # Python 3.9+

save_folder = os.path.expanduser("~/Downloads/Hapag_XTs/Hapag_XTs_Yesterday")
os.makedirs(save_folder, exist_ok=True)

central = ZoneInfo("America/Chicago")
now = datetime.now(central) - timedelta(days=0)  # adjust
start_today = now.replace(hour=0, minute=0, second=0, microsecond=0)
end_today = now.replace(hour=23, minute=59, second=59, microsecond=999999)

download_folder = os.path.expanduser("~/Downloads/Hapag_XTs")
print("")
print("Clearing download folder of all its contents...")
if os.path.exists(download_folder):
    shutil.rmtree(download_folder)
    print("Cleared download folder.")
    print("")
    print("")
    os.makedirs(download_folder, exist_ok=True)

print("Now grabbing emails from Inbox")
print("")
time.sleep(2)

########
fetch_and_save_wosd_pdfs('Inbox', start_dt=start_today, end_dt=end_today)
########

print("Now grabbing emails from Actioned")
print("")
time.sleep(2)

########
fetch_and_save_wosd_pdfs('Actioned', start_dt=start_today, end_dt=end_today)
########

# SEARCHING YESTERDAY'S EMAILS (Central time, 3:59PM-11:59PM)
prev_day = now - timedelta(days=1)
start_prev = prev_day.replace(hour=15, minute=59, second=0, microsecond=0)
end_prev = prev_day.replace(hour=23, minute=59, second=59, microsecond=999999)
save_folder = os.path.expanduser("~/Downloads/Hapag_XTs/Hapag_XTs_Yesterday")
os.makedirs(save_folder, exist_ok=True)

print("Now grabbing emails from Actioned for yesterday from 3:59 PM to 11:59 PM")
print("")
time.sleep(2)

#########
fetch_and_save_wosd_pdfs('Actioned', save_folder=save_folder, start_dt=start_prev, end_dt=end_prev)
#########

clean_xts()

send_email()
print("Closing in 3 seconds")
time.sleep(3)