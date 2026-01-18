import PyPDF2
import docx
import numpy as np
import os
import re
import json
import language_tool_python as llp
from sklearn.feature_extraction.text import TfidfVectorizer
import pandas as pd
import math

# set maximum points for each grading category
UNIFORMITY_WEIGHT = 6.0
PURPOSE_WEIGHT = 4.0
SPELLING_AND_GRAMMAR_WEIGHT = 3.0
LENGTH_PAGES_WEIGHT = 2.0

def calculate_uniformity(content):
    # measures how much sentence lengths vary for better reading flow
    sentences = re.split(r'[.?!]+', content)
    sentences = [s.strip() for s in sentences if len(s.split()) > 3]
    if len(sentences) < 5: return 0
    lengths = [len(s.split()) for s in sentences]
    std_dev = np.std(lengths)
    target_std = 10
    score = max(0, UNIFORMITY_WEIGHT - (abs(std_dev - target_std) * 0.3))
    return round(score, 2)

def calculate_purpose(content):
    # uses tf-idf to find important keywords and check topic depth
    if len(content.split()) < 50: return 0.0
    try:
        vectorizer = TfidfVectorizer(stop_words='english', max_features=50)
        matrix = vectorizer.fit_transform([content])
        feature_names = vectorizer.get_feature_names_out()
        vocabulary_score = (len(feature_names) / 50) * PURPOSE_WEIGHT
        return round(min(PURPOSE_WEIGHT, vocabulary_score), 2)
    except:
        return 0.0

def calculate_grammar(content, tool):
    # calculates score based on number of errors per hundred words
    matches = tool.check(content)
    total_words = len(content.split())
    if total_words == 0: return 0
    error_rate = (len(matches) / total_words) * 100
    score = max(0, SPELLING_AND_GRAMMAR_WEIGHT - (error_rate * 0.5))
    return round(score, 2)

def extract_text(file_path):
    # extracts text content from pdf or docx files automatically
    ext = os.path.splitext(file_path)[1].lower()
    text = ""
    try:
        if ext == ".pdf":
            reader = PyPDF2.PdfReader(file_path)
            pages = reader.pages[5:-1] if len(reader.pages) > 5 else reader.pages
            text = "\n".join([p.extract_text() for p in pages if p.extract_text()])
        elif ext == ".docx":
            doc = docx.Document(file_path)
            text = "\n".join([para.text for para in doc.paragraphs])
    except Exception as e:
        print(f"error reading {file_path}: {e}")
    return text

# load group data from the external json file
json_path = 'groups.json' 
with open(json_path, 'r') as f:
    group_data = json.load(f)

# initialize grammar tool and folder path
tool = llp.LanguageTool('en-US')
folder_path = '../casestudy_1'
export_data = []

if os.path.exists(folder_path):
    # get list of all documents in the folder
    files = [f for f in os.listdir(folder_path) if f.endswith(('.pdf', '.docx'))]
    for file_name in files:
        print(f"grading file: {file_name}")
        full_path = os.path.join(folder_path, file_name)
        content = extract_text(full_path)
        
        # find the group number within the filename
        group_match = re.search(r'\d+', file_name)
        group_id = group_match.group() if group_match else None
        
        # calculate all individual rubric scores
        s_uni = calculate_uniformity(content)
        s_pur = calculate_purpose(content)
        s_gra = calculate_grammar(content, tool)
        s_len = min(LENGTH_PAGES_WEIGHT, (len(content.split()) / 1000) * LENGTH_PAGES_WEIGHT)
        total = math.ceil(s_uni + s_pur + s_gra + s_len)
        
        # pull matching info from the json data
        info = group_data.get(str(group_id), {"topic": "Unknown", "students": []})
        
        # create a separate entry for every student in the group
        for student in info.get("students", []):
            export_data.append({
                "Student Name": student,
                "Group ID": int(group_id) if group_id else 0,
                "Topic": info.get("topic", "N/A"),
                "Sentence Uniformity": s_uni,
                "Clear Purpose": s_pur,
                "Grammar & Spelling": s_gra,
                "Length & Format": round(s_len, 2),
                "Total Score / 15": total
            })

    # sort the data by group id and save to excel
    df = pd.DataFrame(export_data)
    df = df.sort_values(by=["Group ID", "Student Name"])
    output_name = folder_path + "_individual_grades.xlsx"
    df.to_excel(output_name, index=False)
    print(f"done individual grades saved to {output_name}")
else:
    print("Folder Not Found")