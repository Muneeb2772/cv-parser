from flask import Flask, render_template, request, redirect, url_for, send_file, session
import os
import re
import csv
from io import StringIO
from pdfminer.high_level import extract_text
from docx import Document

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # Needed to use session

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

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

def parse_resumes(files):
    parsed_data = []
    for file in files:
        filename = file.filename
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(file_path)

        if filename.endswith('.pdf'):
            text = extract_text_from_pdf(file_path)
        elif filename.endswith('.docx'):
            text = extract_text_from_docx(file_path)
        else:
            continue

        info = extract_info(text)
        parsed_data.append({'Filename': filename, 'Name': info['Name'], 'Email': info['Email']})

    return parsed_data

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    if 'resumes' not in request.files:
        return redirect(request.url)

    files = request.files.getlist('resumes')

    if len(files) == 0:
        return redirect(request.url)

    # Parse the uploaded resumes
    parsed_resumes = parse_resumes(files)

    # Store parsed data in session
    session['parsed_resumes'] = parsed_resumes

    return render_template('results.html', resumes=parsed_resumes)

@app.route('/export_csv')
def export_csv():
    parsed_resumes = session.get('parsed_resumes', [])
    
    # Generate CSV data
    si = StringIO()
    csv_writer = csv.DictWriter(si, fieldnames=['Filename', 'Name', 'Email'])
    csv_writer.writeheader()
    csv_writer.writerows(parsed_resumes)
    output = si.getvalue()
    si.close()

    return send_file(
        StringIO(output),
        mimetype='text/csv',
        attachment_filename='parsed_resumes.csv',
        as_attachment=True
    )

if __name__ == '__main__':
    app.run(debug=True)
