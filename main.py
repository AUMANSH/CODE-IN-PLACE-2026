import datetime
import os
import glob
import pandas as pd
import parser
import matcher
import outlook_manager
import native_dialogs
import ctypes

def show_message_box(title, message):
    ctypes.windll.user32.MessageBoxW(0, message, title, 0)

# Global cursor for scheduling
LAST_SCHEDULED_DT = None

def get_next_available_slots(outlook_client, start_from_dt, num_slots=3, diverse_days=False):
    """
    Generates 'num_slots' available 1-hour slots starting from 'start_from_dt'.
    Uses the provided outlook_client to check availability.
    If diverse_days is True, slots will be picked from different days.
    """
    slots = []
    
    if start_from_dt.minute > 0 or start_from_dt.second > 0:
        start_from_dt = start_from_dt.replace(minute=0, second=0, microsecond=0) + datetime.timedelta(hours=1)
        
    current_dt = start_from_dt
    checks = 0
    MAX_CHECKS = 1000 
    
    while len(slots) < num_slots and checks < MAX_CHECKS:
        checks += 1
        
        # 1. Constraints
        if current_dt.weekday() == 6: # Sunday
            days_to_add = 1
            current_dt = current_dt + datetime.timedelta(days=days_to_add)
            current_dt = current_dt.replace(hour=9, minute=0)
            continue
            
        if current_dt.hour < 9:
            current_dt = current_dt.replace(hour=9, minute=0)
        elif current_dt.hour >= 17:
            current_dt = current_dt + datetime.timedelta(days=1)
            current_dt = current_dt.replace(hour=9, minute=0)
            continue
            
        # 2. Availability
        is_free = True
        if outlook_client and outlook_client.outlook:
            is_free = outlook_client.is_time_free(current_dt)
        
        if is_free:
            end_t = current_dt + datetime.timedelta(hours=1)
            display_str = current_dt.strftime("%A, %b %d") + f" at {current_dt.strftime('%I:%M %p')} - {end_t.strftime('%I:%M %p')}"
            slots.append({
                "start_dt": current_dt,
                "display": display_str
            })
            
            if diverse_days:
                # Move to next day 9 AM for the next slot search
                current_dt = current_dt + datetime.timedelta(days=1)
                current_dt = current_dt.replace(hour=9, minute=0)
                continue
                
        current_dt += datetime.timedelta(hours=1)
        
    return slots

def load_resumes_from_csvs(directory):
    """
    Scans directory for *_structured.csv files.
    Identifies which ones are Resumes (contain Education/Experience/Skills).
    Returns a list of dicts: {'filename': str, 'text': str, 'email_hint': str}
    """
    candidates = []
    csv_files = glob.glob(os.path.join(directory, "*_structured.csv"))
    
    print(f"Found {len(csv_files)} structured CSV files.")
    
    for file_path in csv_files:
        try:
            df = pd.read_csv(file_path)
            # Check if this is a resume
            # A resume usually has 'Experience', 'Education', or 'Skills' sections.
            unique_sections = df['Section'].unique() if 'Section' in df.columns else []
            is_resume = any(sec in unique_sections for sec in ['Experience', 'Education', 'Skills'])
            
            if not is_resume:
                # Fallback: check if content looks like a resume (keyword heuristic)
                all_text = " ".join(df['Value'].astype(str).tolist()).lower()
                if not any(k in all_text for k in ['curriculum vitae', 'resume', 'education', 'experience']):
                    # Likely Aadhar/Pan/Bank
                    # print(f"Skipping {os.path.basename(file_path)} (Not identified as Resume)")
                    continue

            # Reconstruct text
            # We join 'Value' where 'Sub-Field' is 'Content' or generic text rows
            # We also include contact info to ensure parser finds email
            text_parts = []
            email_val = None
            
            # Extract Text
            if 'Value' in df.columns:
                 text_parts = df['Value'].astype(str).tolist()
            
            # Extract Email Hint if explicitly parsed in CSV
            if 'Sub-Field' in df.columns and 'Value' in df.columns:
                emails = df[df['Sub-Field'] == 'Email']['Value'].tolist()
                if emails:
                    email_val = emails[0]

            full_text = "\n".join(text_parts)
            
            candidates.append({
                "filename": os.path.basename(file_path),
                "text": full_text,
                "email_hint": email_val
            })
            
        except Exception as e:
            print(f"Error reading {os.path.basename(file_path)}: {e}")
            
    return candidates

def main():
    global LAST_SCHEDULED_DT
    LAST_SCHEDULED_DT = datetime.datetime.now() + datetime.timedelta(hours=1)
    
    min_score = 60.0 

    print("Requesting Job Description File...")
    jd_filter = "All Supported\0*.txt;*.docx;*.pdf;*.png;*.jpg;*.jpeg\0Text Files\0*.txt\0Word Docs\0*.docx\0PDF\0*.pdf\0Images\0*.png;*.jpg;*.jpeg\0All Files\0*.*\0"
    jd_file = native_dialogs.ask_open_filename(title="Select Job Description File", filter_str=jd_filter)

    if not jd_file:
        print("No Job Description selected. Exiting.")
        return
    
    try:
        if jd_file.lower().endswith(".pdf"):
            jd_text = parser.extract_text_from_pdf(jd_file)
        elif jd_file.lower().endswith(".docx"):
            jd_text = parser.extract_text_from_docx(jd_file)
        elif jd_file.lower().endswith((".png", ".jpg", ".jpeg")):
            print("Extracting text from Image JD (OCR)...")
            jd_text = parser.extract_text_from_image(jd_file)
        else:
            with open(jd_file, "r") as f: jd_text = f.read()
    except Exception as e:
        print(f"Error reading JD: {e}")
        return

    print(f"Job Description Loaded from: {jd_file}")
    
    # --- Auto-Send Confirmation ---
    print("\n" + "="*50)
    print("           CONFIGURATION MODE")
    print("="*50)
    ans = input("ENABLE AUTO-SEND MODE? (If 'y', emails/invites will be SENT immediately. If 'n', they will DRAFT)\n(y/N): ").strip().lower()
    auto_send_mode = (ans == 'y')
    
    if auto_send_mode:
        print(">> WORKING MODEL: AUTO-SEND ENABLED. <<")
    else:
        print(">> DEMONSTRATION MODE: DRAFTS ONLY. <<")
    print("="*50 + "\n")

    # --- Configuration for CSVs ---
    # This matches the OUTPUT_DIRECTORY in folderscandidate.py (default location)
    CANDIDATE_DATA_DIR = r"C:\Users\gupta\...."
    
    # --- Resume Selection Logic ---
    resumes_to_process = []
    
    # 1. Check for processed CSVs in the candidate data directory
    csv_pattern = os.path.join(CANDIDATE_DATA_DIR, "*_structured.csv")
    csv_exists = glob.glob(csv_pattern)
    use_csv = False
    
    if csv_exists:
        print(f"\nDetected {len(csv_exists)} CSV files in: {CANDIDATE_DATA_DIR}")
        # Ask user nicely
        ans = input("Do you want to process these pre-scanned candidates? (Y/n): ").strip().lower()
        if ans != 'n':
            use_csv = True
            print("Loading resumes from CSVs...")
            resumes_to_process = load_resumes_from_csvs(CANDIDATE_DATA_DIR)
            print(f"Loaded {len(resumes_to_process)} resumes from CSV data.")
    
    # 2. Fallback to Manual Selection if No CSV used or No resumes found in CSVs
    if not use_csv or not resumes_to_process:
        if use_csv:
            print("No valid resumes found in known CSVs. Reverting to manual selection.")
            
        print("\nRequesting Resume Files (Manual Selection)...")
        resume_filter = "Documents\0*.pdf;*.docx\0All Files\0*.*\0"
        resume_files = native_dialogs.ask_open_filenames(title="Select Resume Files (PDF/DOCX)", filter_str=resume_filter)

        if not resume_files:
            print("No resumes selected. Exiting.")
            return

        print(f"Selected {len(resume_files)} resumes.")
        
        # Load text from manual files
        for resume_path in resume_files:
            try:
                resume_text = ""
                if resume_path.lower().endswith(".pdf"):
                    resume_text = parser.extract_text_from_pdf(resume_path)
                elif resume_path.lower().endswith(".docx"):
                    resume_text = parser.extract_text_from_docx(resume_path)
                else:
                    with open(resume_path, "r", errors='ignore') as f: resume_text = f.read()
                
                if not resume_text or not resume_text.strip():
                    print(f"Warning: Empty text for {resume_path}")
                    continue
                    
                resumes_to_process.append({
                    "filename": os.path.basename(resume_path),
                    "text": resume_text,
                    "email_hint": None
                })
            except Exception as e:
                print(f"Error reading resume {resume_path}: {e}")

    # 3. Robust Ranking Loop
    print("\nRanking resumes...")
    ranked_candidates = []
    
    for cand in resumes_to_process:
        try:
            resume_text = cand['text']
            filename = cand['filename']
            email_hint = cand['email_hint']
            
            match_result = matcher.match_skills(resume_text, jd_text)
            
            name, email = parser.extract_candidate_info(resume_text)
            
            # Prefer explicit email from CSV if text extraction failed
            if not email and email_hint:
                email = email_hint
            
            if not name: 
                # Try cleaning filename
                name = filename.replace("_structured.csv", "").replace(".pdf", "")
            
            ranked_candidates.append({
                "name": name,
                "email": email,
                "score": match_result["score"],
                "missing": match_result["missing_keywords"],
                "filename": filename
            })
        except Exception as e:
            print(f"Error ranking {cand.get('filename')}: {e}")

    ranked_candidates.sort(key=lambda x: x['score'], reverse=True)
    
    summary_text = ""
    for c in ranked_candidates:
        summary_text += f"{c['filename']}: {c['score']}% (Missing: {', '.join(c['missing'][:5])})\n"
    
    print("\n--- Ranking Results ---")
    print(summary_text)
    show_message_box("Ranking Results", summary_text)
    
    # 4. Process Candidates
    print(f"\nProcessing all candidates (Threshold: {min_score}%)...")
    
    outlook = outlook_manager.OutlookClient()
    if not outlook.outlook:
        print("Could not connect to Outlook. Exiting.")
        return

    for cand in ranked_candidates:
        score = cand['score']
        candidate_name = cand['name']
        candidate_email = cand['email']
        
        if not candidate_email:
            print(f"Skipping {candidate_name} (No email found)")
            continue
        
        print(f"\nProcessing: {candidate_name} ({candidate_email}) - Score: {score}%")
        
        if score < min_score:
            # Decline Logic
            missing_keywords = cand['missing']
            reason_sentence = "While your background is impressive, we are looking for candidates with a closer match to our specific requirements at this time."
            if missing_keywords:
                s_str = ", ".join(missing_keywords)
                reason_sentence = f"Specifically, the role requires experience with {s_str}."
            
            subject = f"Update on your application - {candidate_name}"
            body = f"Dear {candidate_name},\n\n" \
                   f"Thank you for your interest in the position. We appreciate the time you took to apply.\n\n" \
                   f"{reason_sentence}\n\n" \
                   f"We will keep your resume in our database for future opportunities that match your skills.\n\n" \
                   f"Best regards,\nHR Team"
            
            if outlook.send_email(candidate_email, subject, body, auto_send=auto_send_mode):
                action = "SENT" if auto_send_mode else "DRAFTED"
                print(f"-> Decline email {action} for {candidate_name}")
        else:
            # Success Logic
            print("Checking availability for interview slots...")
            # Pass existing outlook client to avoid cache clearing/conflict!
            # Use diverse_days=True to spread options across different days
            slots = get_next_available_slots(outlook, LAST_SCHEDULED_DT, num_slots=3, diverse_days=True)
            
            if not slots:
                 print("WARNING: Could not find any free slots nearby.")
                 continue
                 
            chosen_slot = slots[0]
            LAST_SCHEDULED_DT = chosen_slot['start_dt'] + datetime.timedelta(hours=1)
            start_time_str = chosen_slot['start_dt'].strftime("%Y-%m-%d %H:%M")
            
            subject = f"Interview Invitation - {candidate_name}"
            others_text = ""
            if len(slots) > 1:
                 others_text = "If this time does not work, please reply with one of the alternative slots below:\n"
                 for i, s in enumerate(slots[1:]):
                     others_text += f"{i+1}. {s['display']}\n"
            
            body = f"Dear {candidate_name},\n\nWe were impressed by your profile and would like to invite you for an interview.\n\nWe have scheduled a time for you: {chosen_slot['display']}.\nA Teams meeting link should be attached to this invite.\n\n{others_text}\nWe look forward to speaking with you.\n\nBest,\nHR Team"
            
            if outlook.create_meeting(candidate_email, subject, start_time_str, duration_minutes=60, body=body, auto_send=auto_send_mode):
                action = "SENT" if auto_send_mode else "DRAFTED"
                print(f"-> Meeting invite {action} for {candidate_name} at {start_time_str}")

if __name__ == "__main__":
    main()
