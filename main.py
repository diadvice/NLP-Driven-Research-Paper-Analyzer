import re
import spacy
import PyPDF2
import docx
import os

# Load the small English NLP model
nlp = spacy.load('en_core_web_sm')

def delete_noncharacters(stringg):
    new_str=''
    for i in stringg:
        x=ord(i)
        if x==44 or 65<=x<=90 or 97<=x<=122 or i.isspace():
            new_str+=i
    return new_str

# Extract text from PDF
def extract_text_from_pdf(file_path):
    with open(file_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = ''
        for page in reader.pages:
            text += page.extract_text()
    return text

# Extract text from Word
def extract_text_from_word(file_path):
    doc = docx.Document(file_path)
    text = ''
    for para in doc.paragraphs:
        text += para.text
    return text

# Use NLP to extract people (authors) and dates
def extract_nlp_based_info(text):
    doc = nlp(text)
    authors = set()
    date = None
    
    for ent in doc.ents:
        if ent.label_ == 'PERSON':
            authors.add(ent.text)
        elif ent.label_ == 'DATE' and date is None:  # Take the first date occurrence
            date = ent.text
    
    return list(authors), date

def clean_author_name(author):
    # Use regex to remove any numbers or commas following the author's name
    return re.sub(r'\d+|,', '', author).strip()  # Remove digits and commas

# Define function to extract key info
def extract_info(text):
    info = {}

    # Extract the first few lines of the document
    lines = text.strip().split("\n")
    
    # Assume the first line is the title if it's long enough
    info['name'] = lines[0] if len(lines[0]) > 5 else None  # Set a minimum length for titles
    
    # Use NLP to extract persons from the first few lines to identify authors
    authors = []
    for line in lines[:3]:  # Check the first 3 lines for authors
        doc = nlp(line)
        for ent in doc.ents:
            if ent.label_ == 'PERSON':
                authors.append(ent.text)
    
    # Use the first recognized person entities as authors
    filtered_authors = [clean_author_name(author) for author in authors if author and author.strip()]

    # Join only the valid author names (up to the first 3)
    info['authors'] = ', '.join(filtered_authors[:3]) if filtered_authors else None  # Include only the first 3 authors
    info['authors'] = delete_noncharacters(info['authors'])

    # Use NLP to extract date (check entire text for dates)
    doc = nlp(text)
    date = None
    for ent in doc.ents:
        if ent.label_ == 'DATE':
            date = ent.text
            break  # Take the first occurrence of the date
    
    info['date'] = date

    # Extract conclusion (basic heuristic based on section heading)
    conclusion_match = re.search(r'Conclusion[\s\S]+?(\n\n|\Z)', text, re.IGNORECASE)  # Stop at double newline or end of text
    info['conclusion'] = conclusion_match.group(0) if conclusion_match else None
    if info['conclusion']:
        # Skip the heading
        conclusion_lines = info['conclusion'].split('\n')[1:]  # Skip the heading line
        conclusion_text = ' '.join(conclusion_lines).strip()  # Join the remaining lines
        
        # Split the text into sentences
        sentences = re.split(r'(?<=[.!?]) +', conclusion_text)  # Split by sentence end markers
        filtered_sentences = []
        for sentence in sentences:
            if "Acknowledgments" in sentence or "References" in sentence:
                break 
            filtered_sentences.append(sentence)
        # Include the first 5 sentences
        info['conclusion'] = ' '.join(filtered_sentences[:5])  # Join the first 5 sentences

    # If the paper mentions GitHub, look for a link
    github_match = re.search(r'(https?://github\.com/[^\s]+)', text)
    info['github'] = github_match.group(1) if github_match else None

    # Extract noun chunks and filter out PERSON entities
    keywords = []
    for chunk in doc.noun_chunks:
        if not any(ent.label_ == 'PERSON' for ent in chunk.ents):  # Exclude PERSON entities
            if len(chunk.text.split()) > 1:  # Add chunk to keywords if it's longer than 1 word
                keywords.append(chunk.text)

    # Count keyword frequencies
    keyword_freq = {}
    for keyword in keywords:
        keyword_freq[keyword] = keyword_freq.get(keyword, 0) + 1

    # Sort keywords by frequency
    sorted_keywords = sorted(keyword_freq.items(), key=lambda x: x[1], reverse=True)

    # Filter out articles ('the', 'a', 'an') from the beginning of keywords
    articles = ['the', 'a', 'an']
    filtered_keywords = []
    for word, freq in sorted_keywords:
        word_parts = word.split()
        if word_parts[0].lower() in articles:
            word = ', '.join(word_parts[1:])
        
        # Add the cleaned-up word to the filtered keywords list
        if word.lower() not in ['et al', 'etal']:  # Remove 'et al' and similar terms
            filtered_keywords.append((word, freq))

    # Get the top 5 keywords
    info['topic'] = [kw[0] for kw in filtered_keywords[:5]]  # Get top 5 keywords
    return info

# Google Sheets code stays the same...
import gspread
from oauth2client.service_account import ServiceAccountCredentials

def connect_to_google_sheets(sheet_name):
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(creds)
    sheet = client.open(sheet_name).sheet1
    return sheet

def update_sheet(sheet, info):
    # Limit the text length for fields like 'conclusion'
    max_cell_length = 10000
    if info['conclusion'] and len(info['conclusion']) > max_cell_length:
        info['conclusion'] = info['conclusion'][:max_cell_length] + '...'  # Truncate and add '...'

    # Join the topics into a single string if it exists
    topics = ', '.join(info.get('topic', [])) if isinstance(info.get('topic', []), list) else ''

    # Insert data in the first available row
    row = [
        info.get('name', ''),
        info.get('date', ''),
        info.get('authors', ''),
        info.get('github', ''),
        info.get('conclusion', ''),
        topics  # Add the joined topics here
    ]
    
    sheet.append_row(row)

def process_file(file_path, file_type, sheet_name):
    sheet = connect_to_google_sheets(sheet_name)
    
    # Extract text based on file type
    if file_type == 'pdf':
        text = extract_text_from_pdf(file_path)
    elif file_type == 'docx':
        text = extract_text_from_word(file_path)
    
    # If text extraction fails, skip processing
    if text is None:
        print(f"Skipping {file_path} due to file read error.")
        return
    # Extract key information from text using NLP
    info = extract_info(text)
    
    # Update Google Sheets with the extracted info
    update_sheet(sheet, info)

# New function to process multiple files
def process_multiple_files(num_files, sheet_name):
    for i in range(num_files):
        file_path = input(f"Enter the path for paper {i+1}: ")
        file_type = file_path.split('.')[-1].lower()  # Determine file type based on extension
        if file_type not in ['pdf', 'docx']:
            print(f"Unsupported file type for {file_path}. Skipping...")
            continue
        if not os.path.exists(file_path):
            print(f"File {file_path} does not exist. Skipping...")
            continue
        process_file(file_path, file_type, sheet_name)

# Ask user for the number of files to process
num_files = int(input("Enter the number of papers to process: "))
sheet_name = 'My Research Data Sheet'
process_multiple_files(num_files, sheet_name)
