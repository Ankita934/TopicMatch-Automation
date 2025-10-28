import os
from docx import Document
import pandas as pd
from rapidfuzz import fuzz

# ---------- CONFIG ----------
WORD_FOLDER = "word_files"          # folder containing .docx notes
SYLLABUS_FILE = "syllabus.txt"      # text file with one topic per line
OUTPUT_FILE = "syllabus_check_report.xlsx"
THRESHOLD = 70                      # matching sensitivity (0–100)
# -----------------------------

def extract_text_from_docx(path):
    """Extract text from a Word (.docx) file"""
    doc = Document(path)
    text = "\n".join([para.text for para in doc.paragraphs])
    return text.lower()

def load_all_word_texts(folder):
    """Load text from all Word files"""
    all_texts = {}
    for fname in os.listdir(folder):
        if fname.lower().endswith(".docx"):
            path = os.path.join(folder, fname)
            all_texts[fname] = extract_text_from_docx(path)
    return all_texts

def load_syllabus(file):
    """Read syllabus topics from text file"""
    with open(file, "r", encoding="utf-8") as f:
        topics = [line.strip().lower() for line in f if line.strip()]
    return topics

def find_best_match(topic, all_texts):
    """Find which file matches a syllabus topic best"""
    best_match = {"file": None, "score": 0, "snippet": ""}
    for fname, text in all_texts.items():
        score = fuzz.partial_ratio(topic, text)
        if score > best_match["score"]:
            best_match["score"] = score
            best_match["file"] = fname
            # Try to extract a small snippet for context
            idx = text.find(topic.split()[0])  # look for first word
            if idx != -1:
                start = max(0, idx - 60)
                best_match["snippet"] = text[start: idx + len(topic) + 60].replace("\n", " ")
    return best_match

def main():
    print("Loading syllabus topics...")
    topics = load_syllabus(SYLLABUS_FILE)
    print(f"✅ Loaded {len(topics)} topics")

    print("Reading Word files...")
    all_texts = load_all_word_texts(WORD_FOLDER)
    print(f"✅ Loaded {len(all_texts)} Word files")

    results = []
    for topic in topics:
        match = find_best_match(topic, all_texts)
        found = match["score"] >= THRESHOLD
        results.append({
            "Topic": topic,
            "Found": "Yes" if found else "No",
            "BestScore": match["score"],
            "MatchedFile": match["file"],
            "Snippet": match["snippet"]
        })

    df = pd.DataFrame(results)
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"\n📊 Report saved to: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()