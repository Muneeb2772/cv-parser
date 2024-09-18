import os
import re
import csv
from flask import Flask, render_template, request, redirect, url_for, flash, send_file
from pdfminer.high_level import extract_text
from docx import Document

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'resumes'
app.secret_key = 'your_secret_key'

# Helper functions from your existing backend code
def extract_text_from_pdf(file_path):
    try:
        return extract_text(file_path)
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
        return ""

def extract_text_from_docx(file_path):
    try:
        doc = Document(file_path)
        return '\n'.join([para.text for para in doc.paragraphs])
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
        return ""

def extract_info(text):
    # Patterns for matching names and emails
    name_pattern = r'\b(?:[A-Z][a-z]+(?:\s[A-Z][a-z]+){1,2})\b'
    email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    
    # Look for 'Name:' field explicitly
    name_field_pattern = r'Name:\s*([A-Z][a-zA-Z\s]+)'
    email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    
    # Extracting email addresses
    emails = re.findall(email_pattern, text)
    
    # Check for explicit 'Name:' field
    name_field_match = re.search(name_field_pattern, text)
    if name_field_match:
        name = name_field_match.group(1).strip()
    else:
        # If not found, use general name extraction
        limited_text = text[:500]
        name_matches = re.findall(name_pattern, limited_text)
        filtered_names = [name for name in name_matches if not re.search(excluded_keywords, name, re.IGNORECASE)]
        name = filtered_names[0].strip() if filtered_names else 'N/A'
    
    # Format the email addresses
    email = ', '.join(emails) if emails else 'N/A'
    
    return {'Name': name, 'Email': email}


def parse_resumes(folder_path):
    resumes = []
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        if filename.endswith('.pdf'):
            text = extract_text_from_pdf(file_path)
        elif filename.endswith('.docx'):
            text = extract_text_from_docx(file_path)
        else:
            continue
        info = extract_info(text)
        resumes.append({'Filename': filename, **info})
    return resumes

def write_to_csv(resumes, output_file):
    with open(output_file, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.DictWriter(file, fieldnames=['Filename', 'Name', 'Email'])
        writer.writeheader()
        for resume in resumes:
            writer.writerow(resume)

# Flask routes
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    if 'files[]' not in request.files:
        flash('No file part')
        return redirect(request.url)

    files = request.files.getlist('files[]')

    if not os.path.exists(app.config['UPLOAD_FOLDER']):
        os.makedirs(app.config['UPLOAD_FOLDER'])

    for file in files:
        if file and (file.filename.endswith('.pdf') or file.filename.endswith('.docx')):
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(file_path)

    # Parse the uploaded resumes
    resumes = parse_resumes(app.config['UPLOAD_FOLDER'])

    # Write results to CSV
    output_file = 'output/parsed_resumes.csv'
    if not os.path.exists('output'):
        os.makedirs('output')
    write_to_csv(resumes, output_file)

    # Render the results page with parsed resumes
    return render_template('results.html', resumes=resumes)

@app.route('/download')
def download_csv():
    try:
        return send_file('output/parsed_resumes.csv', as_attachment=True)
    except Exception as e:
        return str(e)