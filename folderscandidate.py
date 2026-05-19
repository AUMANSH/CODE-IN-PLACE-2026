import os
import re
import cv2
import fitz  # PyMuPDF
import pytesseract
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog
from PIL import Image
from io import BytesIO

# --- CONFIGURATION ---
# 1. CENTRAL OUTPUT DIRECTORY
OUTPUT_DIRECTORY = r"C:\Users\gupta\OneDrive - DYNE TECH AI PRIVATE LIMITED\Automated Interview Scheduling System - Version 2"

if not os.path.exists(OUTPUT_DIRECTORY):
    os.makedirs(OUTPUT_DIRECTORY)

# 2. TESSERACT PATH
TESSERACT_CMD = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

if os.path.exists(TESSERACT_CMD):
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD
else:
    print(f"[WARNING] Tesseract not found at {TESSERACT_CMD}. Image PDFs will fail.")

# --- KEYWORDS (Logic Unchanged) ---
RESUME_SECTIONS = {
    "Experience": ["experience", "employment", "work history", "career history"],
    "Education": ["education", "qualifications", "degrees", "certifications"],
    "Skills": ["skills", "technologies", "technical skills", "core competencies"],
    "Projects": ["projects", "key projects", "academic projects"],
    "Summary": ["summary", "profile", "objective", "professional summary"],
    "Contact": ["contact", "address", "phone", "email", "linkedin"]
}

class UniversalPDFExtractor:
    def __init__(self, filepath):
        self.filepath = filepath
        self.filename = os.path.basename(filepath)
        self.directory = os.path.dirname(filepath)
        self.parent_folder_name = os.path.basename(self.directory)

    def preprocess_image_for_ocr(self, image):
        """
        Crucial for Image PDFs: Converts to Black & White to remove background noise.
        """
        img = np.array(image)
        if len(img.shape) == 3:
            img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        gray = cv2.resize(gray, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
        # Thresholding makes text black and background white
        _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        return Image.fromarray(binary)

    def extract_text(self):
        print(f"Processing: {self.filename}")
        try:
            doc = fitz.open(self.filepath)
            full_text = []
            
            for page in doc:
                text = page.get_text()
                # If text is suspiciously short (<50 chars), assume it's an image scan
                if len(text.strip()) < 50:
                    pix = page.get_pixmap(dpi=300)
                    img = Image.open(BytesIO(pix.tobytes("png")))
                    processed_img = self.preprocess_image_for_ocr(img)
                    text = pytesseract.image_to_string(processed_img, lang='eng')
                full_text.append(text)

            return "\n".join(full_text)
        except Exception as e:
            print(f"Error extracting text: {e}")
            return ""

    def classify_document(self, text):
        text_lower = text.lower()
        if ("aadhaar" in text_lower and "dob" in text_lower):
            return "AADHAR"
        elif ("income tax department" in text_lower):
            return "PAN"
        elif ("account number" in text_lower) and ("bank" in text_lower or "ifsc" in text_lower):
            return "BANK_DETAILS"
        elif any(k in text_lower for k in ["experience", "education", "skills"]):
            return "RESUME"
        return "UNKNOWN"

    # --- PARSING LOGIC (PRESERVED) ---
    
    def parse_aadhar(self, text):
        data = []
        section = "Identity Details"
        
        # Aadhar Number (XXXX XXXX XXXX)
        match = re.search(r'\b\d{4}\s\d{4}\s\d{4}\b', text)
        if match: data.append([section, "Aadhar Number", match.group(0)])

        # DOB
        match = re.search(r"DOB\s*[:|-]?\s*(\d{2}/\d{2}/\d{4})", text, re.IGNORECASE)
        if match: data.append([section, "Date of Birth", match.group(1)])

        # Gender
        match = re.search(r"\b(Male|Female|Transgender)\b", text, re.IGNORECASE)
        if match: data.append([section, "Gender", match.group(0)])

        # Name Heuristic
        lines = [l.strip() for l in text.split('\n') if l.strip()]
        if "To" in lines:
            idx = lines.index("To")
            for i in range(1, 4):
                if idx + i < len(lines):
                    cand = lines[idx+i]
                    if re.match(r"^[A-Za-z\s]+$", cand) and len(cand) > 3:
                        data.append([section, "Name", cand])
                        break
        return data

    def parse_pan(self, text):
        data = []
        section = "Identity Details"

        match = re.search(r"[A-Z]{5}[0-9]{4}[A-Z]", text)
        if match: data.append([section, "PAN Number", match.group(0)])

        match = re.search(r"(\d{2}/\d{2}/\d{4})", text)
        if match: data.append([section, "Date of Birth", match.group(1)])

        lines = [l.strip() for l in text.split('\n') if l.strip()]
        for i, line in enumerate(lines):
            if "Name" == line and i+1 < len(lines):
                data.append([section, "Name", lines[i+1]])
                break
            if line.isupper() and "GOVT" not in line and "INCOME" not in line and len(line) > 5 and " " in line:
                if not any(x in line for x in ["FATHER", "DATE", "SIGNATURE"]):
                    data.append([section, "Name Candidate", line])
        return data

    def parse_bank(self, text):
        data = []
        section = "Bank Details"
        
        # 1. Account Number (More flexible regex)
        # Matches "Account Number : 12345" or "Account Number - 12345"
        match = re.search(r"Account\s*Number\s*[:\-\.]?\s*(\d+)", text, re.IGNORECASE)
        if match:
            data.append([section, "Account Number", match.group(1)])

        # 2. IFSC Code (Direct Pattern Search)
        # Looks for 4 letters, 0, then 6 alphanumeric (e.g., HDFC0006552)
        match = re.search(r"\b[A-Z]{4}0[A-Z0-9]{6}\b", text)
        if match:
            data.append([section, "IFSC Code", match.group(0)])

        # 3. Customer Name
        match = re.search(r"Customer\s*Name\s*[:\-\.]?\s*([A-Za-z\s]+)", text, re.IGNORECASE)
        if match:
            name_val = match.group(1).strip()
            # Clean up if it accidentally grabbed trailing junk like "Customer Id"
            name_val = re.split(r'\s{2,}|Customer', name_val)[0]
            data.append([section, "Customer Name", name_val])
            
        # 4. Bank Name (Simple check)
        if "HDFC" in text.upper(): data.append([section, "Bank Name", "HDFC Bank"])
        elif "SBI" in text.upper() or "STATE BANK" in text.upper(): data.append([section, "Bank Name", "State Bank of India"])
        elif "ICICI" in text.upper(): data.append([section, "Bank Name", "ICICI Bank"])

        return data

    def parse_resume(self, text):
        data = []
        lines = text.split('\n')
        current_section = "Header"
        for line in lines:
            line = line.strip()
            if not line: continue
            
            for sec, keywords in RESUME_SECTIONS.items():
                if any(k in line.lower() for k in keywords) and len(line) < 40:
                    current_section = sec
                    break
            
            emails = re.findall(r'[\w\.-]+@[\w\.-]+', line)
            for email in emails: data.append(["Contact", "Email", email])
            
            phones = re.findall(r'(\+?\d[\d\s-]{9,}\d)', line)
            for phone in phones: 
                if len(re.sub(r'\D', '', phone)) >= 10: data.append(["Contact", "Phone", phone])

            if line and "Email" not in line and "Phone" not in line:
                data.append([current_section, "Content", line])
        return data

    def process(self):
        raw_text = self.extract_text()
        if not raw_text: return

        doc_type = self.classify_document(raw_text)
        print(f"  -> Identified: {doc_type}")

        if doc_type == "AADHAR": structured_data = self.parse_aadhar(raw_text)
        elif doc_type == "PAN": structured_data = self.parse_pan(raw_text)
        elif doc_type == "BANK_DETAILS": structured_data = self.parse_bank(raw_text)
        elif doc_type == "RESUME": structured_data = self.parse_resume(raw_text)
        else: structured_data = [["Generic", "Content", raw_text]]

        if structured_data:
            df = pd.DataFrame(structured_data, columns=["Section", "Sub-Field", "Value"])
            
            # --- NAMING & SAVING ---
            # Save format: {ParentFolder}_{OriginalFileName}_structured.csv
            # This ensures uniqueness even if files have the same name in different folders.
            clean_filename = os.path.splitext(self.filename)[0]
            output_filename = f"{self.parent_folder_name}_{clean_filename}_structured.csv"
            output_file = os.path.join(OUTPUT_DIRECTORY, output_filename)
            
            df.to_csv(output_file, index=False)
            print(f"  -> Saved to: {output_file}")
            print("-" * 50)
        else:
            print("  -> No structured data found.")

def main():
    root = tk.Tk()
    root.withdraw()
    
    print("Please select the MAIN FOLDER that contains your Candidate folders.")
    
    # Selecting the Root Directory allows "Multiple Folders" (subfolders) to be processed at once.
    root_folder = filedialog.askdirectory(title="Select Main/Root Folder")
    
    if not root_folder:
        print("No folder selected.")
        input("Press Enter to exit...")
        return

    print(f"\nScanning root: {root_folder}")
    print("="*60)

    pdf_count = 0
    # Walk through the root folder and ALL sub-folders automatically
    for dirpath, dirnames, filenames in os.walk(root_folder):
        for file in filenames:
            if file.lower().endswith(".pdf"):
                pdf_count += 1
                full_path = os.path.join(dirpath, file)
                UniversalPDFExtractor(full_path).process()

    print(f"\nProcessing Complete. {pdf_count} files processed.")
    print(f"All files saved to: {OUTPUT_DIRECTORY}")
    input("Press Enter to exit...")

if __name__ == "__main__":
    main()