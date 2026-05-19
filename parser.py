import fitz  # PyMuPDF
from docx import Document
import os
import re
import spacy
import pytesseract
from PIL import Image

# Configure Tesseract Path
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


# Load Spacy model for name extraction
try:
    nlp = spacy.load("en_core_web_sm")
except OSError:
    print("Warning: Spacy model 'en_core_web_sm' not found. Name extraction may fail.")
    nlp = None

def extract_email(text):
    """Extracts the first email address found in the text."""
    email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    match = re.search(email_pattern, text)
    return match.group(0) if match else None

def extract_name(text):
    """
    Attempts to extract the candidate's name.
    """
    if not text:
        return None
    
    # Pre-cleaning: Remove empty lines from start
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    if not lines:
        return None

    # Look at the first few lines (e.g. up to 5) to find a candidate line that isn't just a page number
    # or junk artifact.
    candidate_line = None
    
    for line in lines[:5]:
        # Regex check: Is it just digits or symbols?
        # e.g. "1", "PAGE 1", "---", etc.
        if re.match(r'^[\d\W]+$', line) or len(line) < 3:
            continue
            
        candidate_line = line
        break
    
    if not candidate_line:
        # Fallback to absolute first line if everything looked "bad"
        candidate_line = lines[0]

    # Strategy 1: Clean excessive titles from the candidate line
    title_pattern = r"^(prof\.?|dr\.?|mr\.?|ms\.?|mrs\.?|er\.?|ar\.?)\s*(\(.*\))?\s*"
    
    cleaned_candidate = re.sub(title_pattern, "", candidate_line, flags=re.IGNORECASE).strip()
    
    # Check validity again after cleaning
    # Must have some letters
    if not re.search(r'[a-zA-Z]', cleaned_candidate):
         return None

    if 1 <= len(cleaned_candidate.split()) <= 7:
        return cleaned_candidate

    # Strategy 2: Spacy NER
    if nlp:
        doc = nlp(text[:2000]) 
        for ent in doc.ents:
            if ent.label_ == "PERSON":
                cleaned_name = ent.text.strip().replace('\n', ' ')
                return cleaned_name
                
    return None

def extract_text_from_image(filepath):
    """Extracts text from an image file using OCR."""
    text = ""
    try:
        image = Image.open(filepath)
        text = pytesseract.image_to_string(image)
    except Exception as e:
        print(f"Error reading Image {filepath}: {e}")
    return text

def extract_text_from_pdf(filepath):
    """Extracts text from a PDF file."""
    text = ""
    try:
        with fitz.open(filepath) as doc:
            for page in doc:
                text += page.get_text()
    except Exception as e:
        print(f"Error reading PDF {filepath}: {e}")
    return text

def extract_text_from_docx(filepath):
    """Extracts text from a DOCX file."""
    text = ""
    try:
        doc = Document(filepath)
        for para in doc.paragraphs:
            text += para.text + "\n"
    except Exception as e:
        print(f"Error reading DOCX {filepath}: {e}")
    return text

def load_resumes(directory):
    """
    Loads all PDF and DOCX resumes from a directory.
    Returns a dict {filename: text}
    """
    resumes = {}
    if not os.path.exists(directory):
        print(f"Directory {directory} not found.")
        return resumes

    for filename in os.listdir(directory):
        filepath = os.path.join(directory, filename)
        if filename.lower().endswith(".pdf"):
            resumes[filename] = extract_text_from_pdf(filepath)
        elif filename.lower().endswith(".docx"):
            resumes[filename] = extract_text_from_docx(filepath)
    return resumes

def extract_candidate_info(text):
    """
    Wrapper to extract both name and email.
    Returns (name, email) tuple.
    """
    return extract_name(text), extract_email(text)
