# CODE-IN-PLACE-2026 HARMONY
HARMONY (Human Resource Automation and Management System) is an integrated human resource management platform designed to streamline and optimize workforce operations. It supports essential HR functions including recruitment, hiring, interview scheduling, employee management, training coordination, and workforce administration while ensuring efficient communication and compliance. The system aims to create a productive, engaged, and high-performing workforce that contributes to the organization’s long-term success.

Prerequisites
1.	Python 3.9+ installed
2.	Tesseract OCR installed
o	Locate it in the project zip or folder:
Tesseract-OCR\tesseract.exe
o	Install it on your system.
3.	Pip available in environment variables

Library Installation
Open the project folder.
Open a new terminal inside it.
Run:
pip install -r requirements.txt
Wait until all libraries are installed successfully.

 Mandatory Configuration
1. Configure Tesseract Path
Open folderscandidate.py
Find in line 21:
TESSERACT_CMD = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
Change it to the location where Tesseract after installing it on your system.

2. Configure Output Directory
In folderscandidate.py, update line 15  :
OUTPUT_DIRECTORY = r"C:\Users\gupta\OneDrive - DYNE TECH AI PRIVATE LIMITED\Automated Interview Scheduling System - Version 2"
Set this to the folder where you want processed candidate data and CSV files saved.

3. Configure Candidate Database Path
Open main.py
Update line around 173:
CANDIDATE_DATA_DIR = r"C:\Users\gupta\OneDrive - DYNE TECH AI PRIVATE LIMITED\Automated Interview Scheduling System - Version 2"
Change it to the folder where your processed candidate database and CSV files are stored.

 Step 1: Candidate Folder Processing
Run:
python folderscandidate.py
The system will ask for the root folder of the candidate.
The script will automatically:
•	Detect PDF types (image-based or text-based)
•	Extract text using OCR if needed
•	Identify document categories:
o	Aadhar
o	PAN
o	Resume
o	Bank details
•	Generate structured outputs
•	Save extracted data into tabular CSV format
All processed files will be saved into the configured output directory.

 Step 2: Resume Screening & Interview Automation
Run:
python main.py
The system will prompt for:
•	Job description file (any format)
•	Resume source:
o	Either a root folder where CSVs are stored
o	Or manual file selection
The system will:
•	Match resumes with job descriptions
•	Score candidates
•	Apply threshold logic
o	Below threshold → rejection email
o	Above threshold → selection flow
You will be asked whether to:
•	Automatically send emails
•	Or only draft them
If selected, the system schedules interviews based on HR availability.

 Step 3: Automated Rescheduling System
Run:
python rescheduler_handler.py
This module:
•	Automatically monitors incoming emails
•	Detects candidate rescheduling requests
•	Validates requested time against HR availability
•	Sends responses automatically
Logic includes:
•	If HR is unavailable → candidate is asked to choose another slot
•	If HR is available → new meeting link is automatically sent
•	No manual mailbox handling is required
This supports multi-layer rescheduling without human intervention.



