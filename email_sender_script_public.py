import pandas as pd
import win32com.client as win32
import os

# Load the Excel file (choose the correct sheet)
EXCEL_FILE = "your_excel_file.xlsx"  # Replace with your file name
while True:
    SHEET_NAME = str(input("Enter the sheet name (e.g., FPPMix or T&M): "))
    if SHEET_NAME not in ["FPPMix", "T&M"]:
        print("Invalid sheet name. Please enter a valid sheet name.")
        continue
    break

df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)

# Read Outlook signature (replace with your signature file name)
signature_path = os.path.join(os.getenv('APPDATA'), 'Microsoft\\Signatures\\YourSignatureFile.htm')
with open(signature_path, 'r') as f:
    signature = f.read()

# Initialize Outlook
outlook = win32.Dispatch('Outlook.Application')

def send_emails(df, signature):
    for index, row in df.iterrows():
        # Unpack row values for clarity
        so_number, project_name, pm_name, email_address, apm_name, email_address2, project_status = row[:7]
        subject = f"Request for updated forecast ETC: SO# {so_number} for {project_name}"
        apm_line = f"{apm_name}," if not isinstance(apm_name, float) else ""
        status_messages = {
            "Yellow": (
                "The project is currently Yellow and needs to be monitored. "
                "We need to try to get it back to Green."
            ),
            "Red": (
                "The project is currently Red and we need to take actions to get it back to Green. "
                "In addition, please provide a deviation comment as to why the project is no longer Green."
            ),
        }
        status_message = status_messages.get(project_status, "")
        status_html = f"{status_message}<br><br>\n" if status_message else ""
        body = (
            f"Hello {pm_name}, {apm_line}<br><br>\n"
            f"Please confirm that the forecast shown in FELIPE for the project is up to date and will be consumed "
            f"for project SO#{so_number} for customer {project_name}. If there will be any changes done to the forecast, "
            f"please provide those so I can update the project.<br><br>\n"
            f"{status_html}"
            f"Once you have provided these updates, I will take the snapshot for the backlog.<br><br>\n"
            f"Regards,<br>\n{signature}\n"
        )
        print("Subject:", subject)
        print("Body:", body)
        print("Recipient:", email_address)
        # Uncomment below to actually send emails
        # mail = outlook.CreateItem(0)
        # mail.Subject = subject
        # mail.To = email_address
        # mail.CC = '' if type(email_address2) == float else email_address2
        # mail.HTMLBody = body
        # mail.Send()
        print(f"Email sent to {email_address} for SO# {so_number}")
    print("All emails sent successfully!")

def test_first_row_email(df, signature):
    """Build and print a test email from the first row of a DataFrame (does not send)."""
    row = df.iloc[0]
    so_number, project_name, pm_name, email_address, apm_name, email_address2, project_status = row[:7]
    subject = f"Request for updated forecast ETC: SO# {so_number} for {project_name}"
    apm_line = f"{apm_name}," if not isinstance(apm_name, float) else ""
    status_messages = {
        "Yellow": (
            "The project is currently Yellow and needs to be monitored. "
            "We need to try to get it back to Green."
        ),
        "Red": (
            "The project is currently Red and we need to take actions to get it back to Green. "
            "In addition, please provide a deviation comment as to why the project is no longer Green."
        ),
    }
    status_message = status_messages.get(project_status, "")
    status_html = f"{status_message}<br><br>\n" if status_message else ""
    body = (
        f"Hello {pm_name}, {apm_line}<br><br>\n"
        f"Please confirm that the forecast shown in FELIPE for the project is up to date and will be consumed "
        f"for project SO#{so_number} for customer {project_name}. If there will be any changes done to the forecast, "
        f"please provide those so I can update the project.<br><br>\n"
        f"{status_html}"
        f"Once you have provided these updates, I will take the snapshot for the backlog.<br><br>\n"
        f"Regards,<br>\n{signature}\n"
    )
    print("Subject:", subject)
    print("Body:", body)
    print("Recipient:", email_address)
    print(f"[TEST] Email would be sent to {email_address} for SO# {so_number}")

if __name__ == "__main__":
    mode = input("Type 'test' to preview the first email, or 'send' to send all emails: ").strip().lower()
    if mode == 'test':
        test_first_row_email(df, signature)
    elif mode == 'send':
        send_emails(df, signature)
    else:
        print("Invalid option. Please type 'test' or 'send'.")
