import PyPDF2
import numpy as np
import os
import re
import language_tool_python as llp
from sklearn.feature_extraction.text import TfidfVectorizer
import pandas as pd
import math

# Criteria & Weighting
#
# Sentence Uniformity:      6.0
# Clear Purpose:            2.0
# Organization:             1.5
# Spelling and Grammar:     2.0
# Tone and Word Choice:     1.5
# Length, Pages, Alighment: 2.0
# Total Marks:             15.0

UNIFORMITY_WEIGHT =           6.0
PURPOSE_WEIGHT =              4.0
SPELLING_AND_GRAMMAR_WEIGHT = 3.0
LENGTH_PAGES_WEIGHT =         2.0


def calculate_uniformity(content):
    sentences = re.split(r'[.?!]+', content)
    sentences = [s.strip() for s in sentences if s.strip()]

    if len(sentences) < 2:
        return 0
    
    lengths = [len(s.split()) for s in sentences]
    standard_deviation = np.std(lengths)

    target_variation = 10
    difference = abs(standard_deviation - target_variation)

    score = max(0, UNIFORMITY_WEIGHT - (difference * 0.1))
    return score

def calculate_purpose(content):
    if not content.strip() or len(content.split()) < 5:
        return 0.0

    try:
        vectorizer = TfidfVectorizer(stop_words='english')
        matrix = vectorizer.fit_transform([content])
        weights = matrix.toarray().flatten()
        
        if weights.size == 0:
            return 0.0

        clarity = np.std(weights) * 500 
        score = (clarity / 100) * PURPOSE_WEIGHT
        return min(PURPOSE_WEIGHT, score)
    except ValueError:
        return 0.0

def calculate_grammar(content):

    matches = tool.check(content)
    total_words = len(content.split())
    
    if total_words == 0:
        return 0

    errors = len(matches)/total_words

    # Prevent Division by zero
    try:
        grammar_and_spelling = max(0, (1 - errors) * SPELLING_AND_GRAMMAR_WEIGHT)
    except ZeroDivisionError:
        grammar_and_spelling = 0
    
    return grammar_and_spelling
    

def calculate_pages_alignment(reader, target_pages=5):
    # get the number of pages found in the pdf
    num_pages = len(reader.pages)
    
    line_indents = []

    # get the various indent values for the pdf
    # Skipping first 2 and last page to check alignment only on body text
    body_pages = reader.pages[2:-1] if num_pages > 3 else reader.pages
    
    for page in body_pages:
        text = page.extract_text()
        if text:
            for line in text.split('\n'):
                if line.strip():
                    line_indents.append(len(line) - len(line.lstrip()))
    
    alignment_max = LENGTH_PAGES_WEIGHT / 2
    page_max = LENGTH_PAGES_WEIGHT / 2

    # calculate alignment score based on the standard deviation of line indents
    alignment_score = max(0, alignment_max - (np.std(line_indents) * 0.25)) if line_indents else 0

    # calculate page_score based on how many pages are missing with a maximum limit
    difference = num_pages - target_pages
    
    if difference < 0:
        page_score = max(0, page_max - (abs(difference) * 0.5))
    else:
        # No punishment for writing more
        page_score = page_max
    
    return round(alignment_score + page_score, 2)


# set up script

# get Files
folder = 'casestudy_1'
entries = os.listdir(folder)
print("All Files: ", entries)
full_paths = [os.path.join(folder, entry) for entry in entries if entry.endswith('.pdf')]

# set up Language Tool
tool = llp.LanguageTool('en-US')

# Mark Storage per File
export_data = []

for i in full_paths:
    print('Grading: ', i)
    total_score = 0.0
    
    # Default rubric to handle failures gracefully
    rubric = {
        "Sentence Uniformity": 0.0,
        "Clear Purpose": 0.0,
        "Spelling and Grammar": 0.0,
        "Length / Pages": 0.0
    }

    try:
        # Open file to read as PDF
        reader = PyPDF2.PdfReader(i)
        num_pages = len(reader.pages)
        full_text = []

        # Logic to skip the first 2 pages (Cover/TOC) and the last page (Bibliography)
        if num_pages > 3:
            content_pages = reader.pages[2:-1]
        else:
            content_pages = reader.pages

        for page in content_pages:
            content_extracted = page.extract_text()
            if content_extracted:
                full_text.append(content_extracted)
            
        content = "\n".join(full_text)

        if content.strip():
            rubric["Sentence Uniformity"] = calculate_uniformity(content)
            rubric["Clear Purpose"] = calculate_purpose(content)
            rubric["Spelling and Grammar"] = calculate_grammar(content)
            rubric["Length / Pages"] = calculate_pages_alignment(reader)
        else:
            rubric["Length / Pages"] = calculate_pages_alignment(reader)

        total_score = math.ceil(sum(rubric.values()))
        print(f"Rubric for {os.path.basename(i)}: {rubric}")

    except Exception as e:
        print(f"Critical error on {i}: {e}")
        total_score = 0.0

    # Add to export list
    export_data.append({
        "File Name": os.path.basename(i),
        "Sentence Uniformity": rubric["Sentence Uniformity"],
        "Clear Purpose": rubric["Clear Purpose"],
        "Spelling and Grammar": rubric["Spelling and Grammar"],
        "Length / Pages": rubric["Length / Pages"],
        "Total Score / 15": total_score
    })
    print("total Score", total_score)

# Export to Excel
df = pd.DataFrame(export_data)
df.to_excel(folder+".xlsx", index=False)

print("\nExport Complete: "+folder+".xlsx")