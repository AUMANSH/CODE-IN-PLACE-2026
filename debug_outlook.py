import win32com.client
import sys
import os

def test_outlook():
    print("--- Outlook Diagnostic ---")
    try:
        print("Attempting to dispatch Outlook.Application...")
        outlook = win32com.client.Dispatch("Outlook.Application")
        print("Dispatch successful.")
        
        print("Attempting to get MAPI namespace...")
        namespace = outlook.GetNamespace("MAPI")
        print("MAPI Namespace acquired.")
        
        print("Attempting to CreateItem(1) [Appointment]...")
        appt = outlook.CreateItem(1)
        print("Appointment item created successfully.")
        
        print("Setting properties...")
        appt.Subject = "Test Appointment from Debug Script"
        appt.Body = "This is a test."
        print("Properties set.")
        
        # Don't save, just verify we got here
        print("Test complete. Outlook integration seems functional.")
        
    except Exception as e:
        print(f"FAILED. Error details: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_outlook()
