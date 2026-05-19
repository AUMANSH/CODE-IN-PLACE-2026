from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import spacy

# Load Spacy model for basic preprocessing (optional but good for cleaning)
try:
    nlp = spacy.load("en_core_web_sm")
except OSError:
    # Fallback if model isn't downloaded yet, though we are installing it.
    # We can try to download it dynamically or just skip advanced NLP for now.
    print("Warning: Spacy model 'en_core_web_sm' not found. Ensure it is installed.")
    nlp = None

def preprocess_text(text):
    """
    Basic cleaning: lowercasing, removing extra whitespace.
    Could be expanded to remove stopwords, lemmatization using Spacy.
    """
    if nlp:
        doc = nlp(text.lower())
        # Example: keep only alpha tokens and remove stopwords
        tokens = [token.lemma_ for token in doc if token.is_alpha and not token.is_stop]
        return " ".join(tokens)
    
    # Fallback simple processing
    return " ".join(text.lower().split())

def calculate_similarity(jd_text, resume_text):
    """
    Calculates cosine similarity between JD and Resume.
    Returns a score (0-100).
    """
    clean_jd = preprocess_text(jd_text)
    clean_resume = preprocess_text(resume_text)

    if not clean_resume or not clean_jd:
        return 0.0

    documents = [clean_jd, clean_resume]
    vectorizer = TfidfVectorizer().fit_transform(documents)
    vectors = vectorizer.toarray()
    
    # Cosine similarity between JD (index 0) and Resume (index 1)
    score = cosine_similarity([vectors[0]], [vectors[1]])[0][0]
    return round(score * 100, 2)

def extract_keywords(text):
    """
    Extracts important keywords (Nouns, Proper Nouns) from text.
    """
    if not nlp:
        return set()
    
    doc = nlp(text.lower())
    # Filter for Nouns and Proper Nouns, remove stopwords
    keywords = {
        token.lemma_ for token in doc 
        if token.pos_ in ["NOUN", "PROPN"] 
        and not token.is_stop 
        and token.is_alpha
    }
    return keywords

def find_missing_keywords(jd_text, resume_text):
    """
    Identifies keywords present in JD but missing in Resume.
    Returns a list of top missing keywords.
    """
    jd_keywords = extract_keywords(jd_text)
    resume_keywords = extract_keywords(resume_text)
    
    missing = list(jd_keywords - resume_keywords)
    # Sort or prioritize? For now, just alphabetize to be deterministic
    missing.sort()
    
    # Return top 5
    return missing[:5]

def rank_resumes(jd_text, resumes_dict):
    """
    Ranks resumes based on similarity to JD.
    Returns a list of tuples: (filename, score, missing_keywords)
    """
    ranked_list = []
    for filename, text in resumes_dict.items():
        score = calculate_similarity(jd_text, text)
        missing_kw = find_missing_keywords(jd_text, text)
        ranked_list.append((filename, score, missing_kw))
    
    ranked_list.sort(key=lambda x: x[1], reverse=True)
    return ranked_list

def match_skills(resume_text, jd_text):
    """
    Matches a single resume against JD text.
    Returns dict: {'score': float, 'missing_keywords': list}
    """
    score = calculate_similarity(jd_text, resume_text)
    missing = find_missing_keywords(jd_text, resume_text)
    return {
        "score": score,
        "missing_keywords": missing
    }
