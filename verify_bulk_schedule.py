from outlook_manager import OutlookClient
import datetime
import time
import main # Import the updated main logic

def verify_bulk_logic():
    print("--- Verifying Bulk Scheduling Logic ---")
    
    # Start checking from 1 hour in the future to ensure we are in a valid window
    start_dt = datetime.datetime.now() + datetime.timedelta(hours=1)
    print(f"Start Time: {start_dt}")
    
    # We call the generator function from main.py
    # Ensure main.py has 'get_next_available_slots' visible
    if not hasattr(main, 'get_next_available_slots'):
        print("ERROR: main.py does not have 'get_next_available_slots' function.")
        return

    slots_1 = main.get_next_available_slots(start_dt, num_slots=3)
    if not slots_1:
        print("FAIL: No slots found (System might be busy or out of work hours).")
        return
        
    print(f"Candidate 1 Options:")
    for s in slots_1:
        print(f"  - {s['display']}")
        
    chosen_1 = slots_1[0]
    print(f"-> Booking Candidate 1 at: {chosen_1['display']}")
    
    # Simulate the logic in main.py: Next search starts 1 hour AFTER the chosen slot starts
    next_cursor = chosen_1['start_dt'] + datetime.timedelta(hours=1)
    
    print(f"Next Cursor: {next_cursor}")
    
    slots_2 = main.get_next_available_slots(next_cursor, num_slots=3)
    if not slots_2:
        print("FAIL: No slots found for Cand 2.")
        return

    print(f"Candidate 2 Options:")
    for s in slots_2:
        print(f"  - {s['display']}")
        
    chosen_2 = slots_2[0]
    
    # Verify sequentiality: Cand 2 must be at least 1 hour after Cand 1
    diff = (chosen_2['start_dt'] - chosen_1['start_dt']).total_seconds()
    if diff >= 3600:
        print(f"PASS: Candidate 2 starts >= 1 hour after Candidate 1 ({diff/3600:.1f} hours diff).")
    else:
        print(f"FAIL: Overlap detected! Diff: {diff} seconds")
        
    # Verify work hours (9 to 17 start times)
    if chosen_1['start_dt'].hour < 9 or chosen_1['start_dt'].hour >= 17:
         print(f"FAIL: Candidate 1 booked at non-work hour {chosen_1['start_dt'].hour}")
         
    if chosen_2['start_dt'].hour < 9 or chosen_2['start_dt'].hour >= 17:
         print(f"FAIL: Candidate 2 booked at non-work hour {chosen_2['start_dt'].hour}")

if __name__ == "__main__":
    try:
        verify_bulk_logic()
    except Exception as e:
        print(f"An error occurred: {e}")
    input("Press Enter to close...") # Keep window open if user runs it
