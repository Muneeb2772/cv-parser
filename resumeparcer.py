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
    excluded_keywords = r'Address|Phone|Location|Email|Contact|Passport|Nationality|Visa|Driving|LinkedIn|Facebook'
    name_pattern = r'([A-Z][a-zA-Z]+(?:\s[A-Z][a-zA-Z]+){0,2})'
    email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    limited_text = text[:500]
    name_match = re.search(name_pattern, limited_text, re.MULTILINE)
    emails = re.findall(email_pattern, limited_text)
    name = name_match.group(1).strip() if name_match else 'N/A'
    if re.search(excluded_keywords, name, re.IGNORECASE):
        name = 'N/A'
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