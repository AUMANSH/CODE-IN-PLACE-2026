
import win32com.client
import time
import datetime

def debug_teams_link_generation():
    print("Starting Outlook Teams Link Debug...")
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        appt = outlook.CreateItem(1) # olAppointmentItem
        
        appt.Subject = "TEST TEAMS LINK GENERATION"
        appt.MeetingStatus = 1 # olMeeting
        
        # Set time to tomorrow
        tomorrow = datetime.datetime.now() + datetime.timedelta(days=1)
        appt.Start = tomorrow.strftime("%Y-%m-%d %H:%M")
        appt.Duration = 30
        
        print("Displaying Appointment...")
        appt.Display()
        
        # Monitor Body/HTMLBody for changes over 10 seconds
        print("Monitoring Body length for 10 seconds...")
        for i in range(10):
            time.sleep(1)
            try:
                # We access these properties to see when/if they get populated
                body_len = len(appt.Body)
                html_body_len = len(appt.HTMLBody)
                print(f"Time {i+1}s: Body Length={body_len}, HTMLBody Length={html_body_len}")
                
                # Check for "Teams" in body
                if "Teams" in appt.Body or "Microsoft Teams" in appt.HTMLBody:
                    print(f" -> FOUND TEAMS LINK at {i+1}s!")
                    break
            except Exception as e:
                print(f"Time {i+1}s: Error accessing properties: {e}")
        
        print("Attempting to access WordEditor...")
        try:
            inspector = appt.GetInspector
            doc = inspector.WordEditor
            print(" -> WordEditor accessed successfully.")
            
            # Try to insert text
            doc.Range(0, 0).InsertBefore("DEBUG PREPENDED TEXT\n\n")
            print(" -> Text prepended via WordEditor.")
            
        except Exception as e:
            print(f" -> FAILED to access WordEditor: {e}")
            
        print("Debug Complete. Please check the open Appointment window for the Teams link and 'DEBUG PREPENDED TEXT'.")
        
        # Don't save or close, just let user look at it.
        # appt.Close(2) # Discard
        
    except Exception as e:
        print(f"FATAL ERROR: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    debug_teams_link_generation()
