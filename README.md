# CODE-IN-PLACE-2026 | HARMONY

## Human Resource Automation and Management System

**HARMONY (Human Resource Automation and Management System)** is an integrated Human Resource platform designed to automate and optimize workforce operations. The system streamlines recruitment, hiring, interview scheduling, employee management, training coordination, and workforce administration while ensuring efficient communication and process compliance

HARMONY helps organizations build a productive, engaged, and high-performing workforce through intelligent automation and decision support.

---

# Features

- Automated candidate document processing
- OCR-based text extraction from PDFs
- Resume screening and candidate scoring
- Automated interview scheduling
- HR availability management
- Email automation for candidate communication
- Automated interview rescheduling
- Structured CSV data generation
- Multi-step recruitment workflow automation

---

# Prerequisites

Before running the project, ensure the following are installed:

- Python 3.9+
- Pip configured in Environment Variables
- Tesseract OCR

### Install Tesseract OCR

Locate the installer inside:

```plaintext
Tesseract-OCR\tesseract.exe
```

Install it on your system.

---

# Installation

Open the project folder and install dependencies:

```bash
pip install -r requirements.txt
```

Wait until all libraries are installed successfully.

---

# Mandatory Configuration

## 1. Configure Tesseract Path

Open:

```plaintext
folderscandidate.py
```

Update:

```python
TESSERACT_CMD = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
```

Change the path to your installed Tesseract location.

---

## 2. Configure Output Directory

Open:

```plaintext
folderscandidate.py
```

Update:

```python
OUTPUT_DIRECTORY = r"C:\Users\YourPath"
```

Set the folder where processed candidate files and CSV outputs should be stored.

---

## 3. Configure Candidate Database Path

Open:

```plaintext
main.py
```

Update:

```python
CANDIDATE_DATA_DIR = r"C:\Users\YourPath"
```

Set the folder containing processed candidate CSV files.

---

# Step 1: Candidate Folder Processing

Run:

```bash
python folderscandidate.py
```

### Workflow

The system prompts for the candidate root folder and automatically:

- Detects PDF type (image-based or text-based)
- Extracts text using OCR
- Identifies document categories:
  - Aadhaar
  - PAN
  - Resume
  - Bank Details
- Generates structured outputs
- Exports extracted data into CSV format

Processed files are saved to the configured output directory.

---

# Step 2: Resume Screening & Interview Automation

Run:

```bash
python main.py
```

### Inputs

- Job description file
- Resume source:
  - Candidate CSV folder
  - Manual file selection

### Processing

The system:

- Matches resumes with job descriptions
- Scores candidates
- Applies threshold logic:
  - Below threshold → Rejection email
  - Above threshold → Selection workflow

### Email Options

- Automatically send emails
- Generate email drafts only

Selected candidates proceed to automated interview scheduling.

---

# Step 3: Automated Rescheduling System

Run:

```bash
python reschedule_handler.py
```

### Functionality

The module automatically:

- Monitors incoming emails
- Detects rescheduling requests
- Checks HR availability
- Sends responses automatically

### Rescheduling Logic

- HR unavailable → Candidate selects another slot
- HR available → New meeting link generated automatically

No manual mailbox handling is required.

---

# Project Structure

```plaintext
CODE-IN-PLACE-2026/
│
├── folderscandidate.py
├── main.py
├── reschedule_handler.py
├── requirements.txt
├── Candidates/
├── outputs/
└── README.md
```

---

# Future Enhancements

- Dashboard Analytics
- AI-based candidate ranking
- Calendar integration
- Multi-user HR support
- Cloud deployment

---

