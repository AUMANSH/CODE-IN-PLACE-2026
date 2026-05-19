from outlook_manager import OutlookClient
import datetime

def verify():
    print("Starting Verification...")
    client = OutlookClient()
    if not client.outlook:
        print("Outlook not available.")
        return

    subject = "VERIFICATION OF FIX"
    start_time = (datetime.datetime.now() + datetime.timedelta(days=1)).strftime("%Y-%m-%d %H:%M")
    body = "This is the custom body text.\nIt should appear ABOVE the Teams link.\nThe Teams link should exist."
    
    print(f"Calling create_meeting for time {start_time}...")
    success = client.create_meeting("test@example.com", subject, start_time, body=body)
    if success:
        print("create_meeting returned True. Check the Outlook window.")
    else:
        print("create_meeting returned False.")

if __name__ == "__main__":
    verify()
