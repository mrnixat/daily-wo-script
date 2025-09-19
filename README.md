# Daily WO Script

This project automates the process of downloading, cleaning, and emailing Hapag Crosstown (XT) work order PDFs from a shared mailbox using Microsoft Graph API, then processes and sends a daily report.

## Features
- Downloads WOSD PDF attachments from specific Outlook mailbox folders (Inbox, Actioned) for today and yesterday.
- Cleans and organizes downloaded PDFs, moving updates to a separate folder.
- Extracts billing references and container numbers from PDFs.
- Generates an Excel report summarizing the work orders and containers.
- Emails the report to specified recipients.

## How It Works
- `RUN.py` orchestrates the workflow:
  1. Clears the download folder.
  2. Downloads WOSD PDFs from Inbox and Actioned folders for today and yesterday.
  3. Cleans and organizes the PDFs.
  4. Generates an Excel report.
  5. Emails the report to recipients.

## Customization
- Change date ranges or mailbox folders by editing `RUN.py` and `vars.py`.
- Update email recipients in `send_email.py`.

