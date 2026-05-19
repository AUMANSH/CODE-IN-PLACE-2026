import win32com.client
import datetime
import spacy
from dateutil import parser as date_parser
import outlook_manager
import os
import sys

# Constants matching main.py
SUBJECT_KEYWORDS = ["Interview Invitation", "Rescheduling Interview"]

class ReschedulingAgent:
    def __init__(self):
        print("Initializing AI Rescheduling Agent...")
        try:
            self.nlp = spacy.load("en_core_web_sm")
            print("-> NLP Model Loaded (spacy: en_core_web_sm)")
        except Exception as e:
            print(f"Error loading NLP model: {e}")
            self.nlp = None

        self.outlook_mgr = outlook_manager.OutlookClient()
        if not self.outlook_mgr.outlook:
            print("Error: Could not connect to Outlook via Manager.")
            sys.exit(1)
            
        self.namespace = self.outlook_mgr.namespace

    def suggest_next_slots(self):
        """Helper to find 3 valid slots starting from tomorrow."""
        # We need to import main locally or access logic to find slots
        # Simpler: just use main's function if possible, or implement simple loop
        import main
        # Start looking from tomorrow 9 AM
        start_search = datetime.datetime.now() + datetime.timedelta(days=1)
        start_search = start_search.replace(hour=9, minute=0, second=0)
        
        slots = main.get_next_available_slots(self.outlook_mgr, start_search, num_slots=3, diverse_days=True)
        text = ""
        for s in slots:
             text += f"- {s['display']}\n"
        return text

    def find_candidate_replies(self):
        """
        Scans Inbox for unread replies to Interview Invitations.
        """
        try:
            inbox = self.namespace.GetDefaultFolder(6) # 6 = Inbox
            messages = inbox.Items
            # Restrict to Unread
            messages = messages.Restrict("[Unread] = True")
            
            replies = []
            for msg in messages:
                try:
                    # Check subject
                    if any(kw in msg.Subject for kw in SUBJECT_KEYWORDS):
                        # Ensure it's a candidate reply (basic check matches Subject)
                        print(f"Found potential reply: {msg.Subject} from {msg.SenderName}")
                        replies.append(msg)
                except Exception as e:
                    print(f"Error checking message: {e}")
                    
            return replies
        except Exception as e:
            print(f"Error scanning inbox: {e}")
            return []

    def parse_time_from_text(self, text):
        """
        Uses NLP and keyword matching to extract the requested date and time from the email body.
        Handles relative times (e.g. '10am' passed today -> tomorrow 10am).
        """
        if not self.nlp:
            return None
            
        doc = self.nlp(text)
        potential_datetimes = []
        now = datetime.datetime.now()
        
        # Keywords to catch time expressions even if Spacy misses them
        time_keywords = ["am", "pm", "morning", "afternoon", "evening", "tomorrow", "next", "monday", "tuesday", "wednesday", "thursday", "friday", "at"]

        for sent in doc.sents:
            # Check 1: Spacy Entities
            has_entity = any(ent.label_ in ["TIME", "DATE"] for ent in sent.ents)
            
            # Check 2: Simple Keyword matching
            sent_lower = sent.text.lower()
            has_keyword = any(k in sent_lower for k in time_keywords)
            
            # Identify if it contains digits (likely time/date)
            has_digits = any(c.isdigit() for c in sent.text)

            if (has_entity or has_keyword) and has_digits:
                try:
                    # Default anchor: Today 9am. 
                    # This helps partial times like "10am" parse correctly.
                    clean_default = now.replace(hour=9, minute=0, second=0, microsecond=0)
                    
                    dt = date_parser.parse(sent.text, fuzzy=True, default=clean_default)
                    
                    if dt < now:
                        # Case 1: Time passed today (e.g. said "9am" at 2pm) -> Assume Tomorrow
                        if dt.date() == now.date():
                            dt = dt + datetime.timedelta(days=1)
                        # Case 2: Past Date (e.g. said "Monday" on Friday -> dateutil gives LAST Monday) -> Assume Next Monday
                        else:
                            # Add 7 days to bring it to the future
                            dt = dt + datetime.timedelta(days=7)
                    
                    # Verify it's future now
                    if dt > now:
                        potential_datetimes.append(dt)
                        # print(f"   [Debug] Found potential time: {dt} in sentence: '{sent.text.strip()}'")
                except:
                    pass
        
        if potential_datetimes:
            # Return the first future date found
            return potential_datetimes[0]
            
        return None

    def clean_reply_body(self, body):
        """
        Removes reply chains, signatures, and headers from the email body
        to isolate the candidate's actual message.
        """
        # Common delimiters for replies
        delimiters = [
            "_____", # Outlook line
            "From:", 
            "On ",   # On [date] wrote:
            "Sent from mine",
            "--"     # Signature
        ]
        
        lines = body.splitlines()
        cleaned_lines = []
        
        for line in lines:
            line = line.strip()
            # Stop if we hit a likely delimiter line
            is_delimiter = False
            for d in delimiters:
                # Relaxed length check for 'On ' headers which can be long
                max_len = 300 if d == "On " else 100
                
                if line.startswith(d) and len(line) < max_len:
                     # Special case for "On ...": ensure it looks like a header (has 'wrote' or a year)
                     if d == "On ":
                         is_header = "wrote" in line or "202" in line or "203" in line
                         if not is_header:
                             continue
                     is_delimiter = True
                     break
            
            if is_delimiter:
                break
                
            cleaned_lines.append(line)
            
        return "\n".join(cleaned_lines)

    def process_reply(self, message):
        """
        Main logic for a single message.
        """
        raw_body = message.Body
        # Clean the body to remove forwarded text so we don't parse the OLD invite time
        body = self.clean_reply_body(raw_body)
        
        print(f"\nProcessing reply from {message.SenderName}...")

        sender_email = ""
        try:
            sender_email = message.SenderEmailAddress
        except:
            print(f"Could not get email for {message.Subject}")
            return
        
        # 1. Parse Time
        requested_dt = self.parse_time_from_text(body)
        
        if not requested_dt:
            print("-> AI could not detect a specific reschedule time in the candidate's reply.")
            return

        print(f"-> Candidate requested time: {requested_dt.strftime('%Y-%m-%d %H:%M')}")
        
        # 2. Check Availability
        print("-> Checking HR availability...")
        
        # FIX: Enforce Business Hours (9 AM - 5 PM) and Weekdays (0-4)
        hour = requested_dt.hour
        is_weekend = requested_dt.weekday() >= 5 # 5=Sat, 6=Sun
        
        if hour < 9 or hour > 17 or is_weekend:
             print(f"-> Requested time {requested_dt.strftime('%H:%M')} is outside business hours (9-5 Mon-Fri).")
             # Send rejection/clarification
             reply = message.Reply()
             reply.Body = (f"Hi {message.SenderName},\n\n"
                           f"The time you requested ({requested_dt.strftime('%A, %b %d at %I:%M %p')}) is outside our business hours.\n"
                           f"Our working hours are Monday to Friday, 09:00 AM to 05:00 PM.\n\n"
                           f"Please propose a time within this window, or select from the text below:\n"
                           f"{self.suggest_next_slots()}")
             reply.Send()
             print("-> Rejection (Business Hours) reply sent to candidate.")
             message.UnRead = False
             return

        is_free = self.outlook_mgr.is_time_free(requested_dt)
        
        reply = message.Reply()
        
        if is_free:
            print("-> Slot is AVAILABLE. Scheduling...")
            
            # 3. Schedule Logic
            # Create new meeting
            subject = f"Rescheduled Interview - {message.SenderName}"
            body_invite = f"Dear {message.SenderName},\n\nAs requested, we have rescheduled your interview to {requested_dt.strftime('%A, %b %d at %I:%M %p')}.\n\nSee you then!"
            
            start_str = requested_dt.strftime("%Y-%m-%d %H:%M")
            success = self.outlook_mgr.create_meeting(sender_email, subject, start_str, duration_minutes=60, body=body_invite, auto_send=True)
            
            if success:
                # Send confirmation reply
                reply.Body = f"Hi {message.SenderName},\n\nI have received your request. I've sent a new calendar invite for {requested_dt.strftime('%I:%M %p')}.\n\nBest,\nHR Team"
                reply.Send()
                print("-> Confirmation reply sent to candidate.")
                
                # Mark original as read
                message.UnRead = False
            else:
                print("-> Failed to create meeting.")
        else:
            print("-> Slot is BUSY.")
            reply.Body = f"Hi {message.SenderName},\n\nUnfortunately, we are not available at {requested_dt.strftime('%I:%M %p')}. Please propose another time."
            reply.Send()
            print("-> Rejection reply sent.")
            message.UnRead = False

    def run(self):
        print("--- AI Rescheduling System ---")
        print("Scanning for 'Interview Invitation' replies...")
        
        replies = self.find_candidate_replies()
        
        if not replies:
            print("No unread replies found.")
            
            # --- DEMO MODE FOR USER ---
            print("\n[DEMO MODE] Since no real emails were found, let's simulate a candidates reply.")
            demo_text = input("Enter a simulated candidate reply (e.g. 'Can we meet next Tuesday at 10am?'): ")
            if not demo_text:
                demo_text = "Actually, I am only free next Friday at 2pm."
            
            print(f"\nSimulating parsing for: '{demo_text}'")
            dt = self.parse_time_from_text(demo_text)
            if dt:
                print(f"AI Parsed Result: {dt}")
                
                # REPLICATE LOGIC FROM process_reply FOR DEMO ACCURACY
                hour = dt.hour
                is_weekend = dt.weekday() >= 5
                
                if hour < 9 or hour > 17 or is_weekend:
                    print("-> [DEMO RESULT] Request REJECTED: Outside Business Hours (9-5 Mon-Fri).")
                    print("   (System would send rejection email with alternative slots)")
                else:
                    is_free = self.outlook_mgr.is_time_free(dt)
                    if is_free:
                        print("-> [DEMO RESULT] Request ACCEPTED: Slot is Available.")
                        print("   (System would book meeting)")
                    else:
                        print("-> [DEMO RESULT] Request REJECTED: Slot is Busy (Calendar Conflict).")
                        
            else:
                print("AI could not parse a date/time from that text.")
            
        else:
            print(f"Found {len(replies)} replies. Processing...")
            for msg in replies:
                self.process_reply(msg)

if __name__ == "__main__":
    agent = ReschedulingAgent()
    agent.run()
