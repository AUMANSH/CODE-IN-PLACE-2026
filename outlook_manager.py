import win32com.client
import datetime
import time
import shutil
import os
import sys
import pythoncom

class OutlookClient:
    def __init__(self):
        self.outlook = None
        try:
            # 1. Clear gen_py cache to prevent "AttributeError" on repeated runs
            # This is a nuclear option but necessary for stability with Outlook COM
            try:
                gen_py_path = os.path.join(os.environ.get('LOCALAPPDATA'), 'Temp', 'gen_py')
                if os.path.exists(gen_py_path):
                    shutil.rmtree(gen_py_path, ignore_errors=True)
            except Exception:
                pass # access denied or minor issue, ignore

            # 2. Initialize COM Library for this thread
            pythoncom.CoInitialize()

            # 3. Use Dispatch
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
        except Exception as e:
            print(f"Error connecting to Outlook: {e}")
            self.outlook = None

    def send_email(self, to_email, subject, body, auto_send=False):
        if not self.outlook:
            print("Outlook not available.")
            return False

        try:
            # Creating item via Application object is safer than Namespace.Items.Add for some versions
            mail = self.outlook.CreateItem(0) # 0 = olMailItem
            mail.To = to_email
            mail.Subject = subject
            mail.Body = body
            
            if auto_send:
                mail.Send()
                print(f"-> Email SENT to {to_email}")
            else:
                mail.Display() 
                print(f"-> Email DRAFTED for {to_email}")
            return True
        except Exception as e:
            print(f"Error creating email: {e}")
            return False

    def is_time_free(self, start_dt, duration_minutes=60):
        """
        Checks if the user is free at the given datetime.
        Returns True if free, False if busy/tentative/OOO.
        """
        if not self.outlook or not self.namespace:
            return True # Fail open if we can't check, but usually safer to assume free? Or fail safe?
            # Let's assume free to proceed, or return False causing skip?
            # Better to assume Free to avoid infinite loops if Outlook is wonky, but print error.
        
        try:
            # 1. Get Default Calendar Folder
            calendar = self.namespace.GetDefaultFolder(9) # 9 = olFolderCalendar
            items = calendar.Items
            items.Sort("[Start]")
            items.IncludeRecurrences = True

            # 2. Define Time Window
            # Format: 'DD/MM/YYYY HH:MM AMPM' or strict generic format usually working is 'YYYY-MM-DD HH:MM' 
            # win32com filters can be tricky with seconds. 
            end_dt = start_dt + datetime.timedelta(minutes=duration_minutes)
            
            # Filter string: "[Start] < 'End' AND [End] > 'Start'"
            # Note: We must format explicitly.
            # Using PyWin32 time formatting can be sensitive to locale.
            # Best safe format often: "MM/DD/YYYY HH:MM" or use Jet format.
            fmt = "%m/%d/%Y %H:%M" 
            s_str = start_dt.strftime(fmt)
            e_str = end_dt.strftime(fmt)
            
            # "Find items that start before our end time AND end after our start time"
            restriction = f"[Start] < '{e_str}' AND [End] > '{s_str}'"
            
            restricted_items = items.Restrict(restriction)
            
            for item in restricted_items:
                # Check BusyStatus
                # 0=Free, 1=Tentative, 2=Busy, 3=OOO, 4=WorkingElsewhere
                if item.BusyStatus != 0: 
                    print(f"  -> Slot {start_dt} blocked by: '{item.Subject}' (Status: {item.BusyStatus})")
                    return False
            
            return True
            
        except Exception as e:
            print(f"Warning: Could not check availability: {e}")
            return True # Fallback to allowing it

    def create_meeting(self, to_email, subject, start_time_str, duration_minutes=60, body="", auto_send=False):
        """
        start_time_str format: 'YYYY-MM-DD HH:MM'
        """
        if not self.outlook:
            print("Outlook not available.")
            return False
        try:
            appt = self.outlook.CreateItem(1) # 1 = olAppointmentItem
            
            # 1. Set recipients and resolve
            appt.Recipients.Add(to_email)
            appt.Recipients.ResolveAll()
            
            # 2. Set Subject, Location, and Meeting Status
            appt.Subject = subject
            appt.Location = "Microsoft Teams Meeting"
            appt.MeetingStatus = 1 # 1 = olMeeting
            appt.Start = start_time_str 
            appt.Duration = duration_minutes

            # 3. Save FIRST, then Display.
            # Saving first ensures the object exists and often helps the Teams Add-in recognize it needs to act.
            appt.Save()
            
            # We must Display to trigger Teams link generation (Add-in requirement)
            # Even if auto-sending, we open it briefly, fill it, then send.
            appt.Display()
            
            # 4. Polling Mechanism: Wait for Teams Link to appear
            # We check via WordEditor because appt.Body/HTMLBody can be flaky or throw errors during generation.
            print("Waiting for Teams link generation...")
            time.sleep(2) # Give a brief moment for window to initialize before hammering it
            link_found = False
            for i in range(20): # Try for up to 20 seconds
                try:
                    inspector = appt.GetInspector
                    doc = inspector.WordEditor
                    current_text = doc.Range().Text
                    
                    # Clean up text for check (remove whitespace sometimes helps)
                    # But broadly looking for key phrases is 100% reliable if they exist.
                    if "Microsoft Teams" in current_text or "Teams Meeting" in current_text:
                        print(f"Teams link detected significantly at {i*1 + 2}s. (Body len: {len(current_text)})")
                        link_found = True
                        break
                    else:
                        if i % 2 == 0: # Log every other second to avoid spam
                            print(f"Polling {i+1}/20s: Link not found yet. Body len: {len(current_text)}")
                        
                        # If body is still size 1 (just newline), the Add-in is likely stuck.
                        # We can try to re-trigger it periodically (every 5 seconds).
                        if i > 0 and i % 5 == 0 and len(current_text) < 5:
                            print(f"Body stuck (len={len(current_text)}). Re-triggering MeetingStatus...")
                            appt.MeetingStatus = 0 # olNonMeeting
                            time.sleep(2.0) # Wait for Outlook to process "off"
                            appt.MeetingStatus = 1 # olMeeting
                            time.sleep(1.0) # Wait for it to turn back on
                            
                except Exception as e:
                    # Ignore COM errors during initialization but log if persistent
                     if i % 2 == 0:
                        print(f"Polling {i+1}/20s: Error reading body: {e}")
                time.sleep(1)
            
            if not link_found:
                print("WARNING: Teams link not detected after 20s. Proceeding anyway, but link might be missing.")

            # 5. Inject custom body SAFELY using WordEditor
            # This preserves the existing body (signature + Teams link)
            try:
                inspector = appt.GetInspector
                doc = inspector.WordEditor
                # Insert at the very beginning
                doc.Range(0, 0).InsertBefore(body + "\n\n")
            except Exception as e:
                print(f"Error injecting body via WordEditor: {e}")
                print("Skipping fallback to avoid overwriting Teams link.")

            # 6. Finalize
            appt.Save()
            
            if auto_send:
                # Use .Send() to send automatically
                appt.Send()
                print(f"-> Meeting Invite SENT to {to_email} at {start_time_str}")
            else:
                # It's already displayed from step 3
                print(f"-> Meeting Invite DRAFTED for {to_email} at {start_time_str}")
                
            return True
        except Exception as e:
            print(f"Error creating meeting: {e}")
            import traceback
            traceback.print_exc()
            return False
