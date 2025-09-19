from datetime import datetime, timedelta
# FOR clean_xts.py and send_email.py
days_back = 0
start = 0
end = 0

# FOR grab_xts.py
start_date = (datetime.now() - timedelta(days=start)) # date to start from
end_date = datetime.now() - timedelta(days=end) # date to stop at