# Automated Outlook Email Sender for Project Forecasts

This script automates the process of sending project forecast update requests via Microsoft Outlook, using data from an Excel file. It is designed for project managers or analysts who need to send personalized emails to multiple recipients based on project status.

## Features
- Reads project and contact data from an Excel file
- Builds personalized email messages for each project
- Option to preview the first email (test mode) or send all emails (send mode)
- Uses your Outlook signature automatically
- Handles special status messages for projects marked as "Yellow" or "Red"

## Requirements
- Python 3.x
- Microsoft Outlook (desktop app, Windows only)
- Python packages: `pandas`, `pywin32`
- An Excel file with the required columns (see below)

## Installation
1. Install Python 3.x if you haven't already.
2. Install required packages:
   ```sh
   pip install pandas pywin32
   ```
3. Place your Excel file in the same directory as the script, or update the `EXCEL_FILE` variable.
4. Update the `signature_path` variable with the correct path to your Outlook signature file (usually in `%APPDATA%\Microsoft\Signatures`).

## Excel File Format
Your Excel file should have a sheet (e.g., `FPPMix` or `T&M`) with the following columns in this order:

1. SO Number
2. Project Name
3. PM Name
4. Email Address
5. APM Name (optional)
6. Email Address 2 (optional)
7. Project Status (e.g., "Green", "Yellow", "Red")

## Usage
Run the script from the command line:

```sh
python email_sender_script_public.py
```

You will be prompted to:
- Enter the sheet name (e.g., `FPPMix` or `T&M`)
- Choose between test mode (preview the first email) or send mode (send all emails)

### Test Mode
- Type `test` when prompted.
- The script will print the first email to the console (no emails are sent).

### Send Mode
- Type `send` when prompted.
- The script will send emails to all recipients in the Excel sheet using Outlook.
- By default, the actual sending lines are commented out for safety. Uncomment them to enable sending.

## Customization
- Update the `EXCEL_FILE` and `signature_path` variables as needed.
- Adjust the email body or subject in the script to fit your needs.

## Disclaimer
- This script is provided as-is. Test thoroughly before using in production.
- Do not share sensitive data (real email addresses, signatures, or confidential project info) in public repositories.

---

Feel free to fork and adapt for your own workflow!
