from flask import Flask, render_template, request, send_file, redirect, url_for
import os
import pandas as pd
import PyPDF2
from werkzeug.utils import secure_filename
import shutil

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'

# Ensure upload and output folders exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    # Clear previous files
    for folder in [app.config['UPLOAD_FOLDER'], app.config['OUTPUT_FOLDER']]:
        for file in os.listdir(folder):
            os.remove(os.path.join(folder, file))

    # Save uploaded files
    pdf_file = request.files['pdf']
    excel_file = request.files.get('excel')  # Excel file is optional
    pages_per_split = int(request.form['pages_per_split'])  # Get the number of pages per split

    pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(pdf_file.filename))
    pdf_file.save(pdf_path)

    if excel_file:
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(excel_file.filename))
        excel_file.save(excel_path)
    else:
        excel_path = None

    # Process the files
    if excel_path:
        process_files(pdf_path, excel_path, pages_per_split)
    else:
        split_pdf(pdf_path, pages_per_split)

    return redirect(url_for('download'))

def split_pdf(pdf_path, pages_per_split):
    with open(pdf_path, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        num_pages = len(reader.pages)

        for i in range(0, num_pages, pages_per_split):
            pdf_writer = PyPDF2.PdfWriter()
            for j in range(i, min(i + pages_per_split, num_pages)):
                pdf_writer.add_page(reader.pages[j])

            output_path = os.path.join(app.config['OUTPUT_FOLDER'], f"split_{i // pages_per_split + 1}.pdf")
            with open(output_path, "wb") as output_pdf:
                pdf_writer.write(output_pdf)

def process_files(pdf_path, excel_path, pages_per_split):
    # Load the Excel sheet
    df = pd.read_excel(excel_path)
    student_names = df['student name'].tolist()

    with open(pdf_path, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        num_pages = len(reader.pages)

        for i in range(0, num_pages, pages_per_split):
            pdf_writer = PyPDF2.PdfWriter()
            for j in range(i, min(i + pages_per_split, num_pages)):
                pdf_writer.add_page(reader.pages[j])

            # Extract text from the first page of the split
            text = reader.pages[i].extract_text()

            # Extract the value of the "Name" tag
            name_tag = "Name"
            name_value = None
            if name_tag in text:
                start_index = text.find(name_tag) + len(name_tag)
                end_index = text.find("Reg No.", start_index)
                if end_index == -1:
                    end_index = len(text)
                name_value = text[start_index:end_index].strip()
            print(f"Extracted name from split {i // pages_per_split + 1}: {text}")

            # Search for the student name in the Excel sheet
            found_name = None
            if name_value:
                for student_name in student_names:
                    if student_name.split(",")[0].strip() in name_value:
                        found_name = student_name
                        break

            # Save the split file with the student name
            if found_name:
                output_path = os.path.join(app.config['OUTPUT_FOLDER'], f"{found_name}.pdf")
            else:
                output_path = os.path.join(app.config['OUTPUT_FOLDER'], f"split_{i // pages_per_split + 1}.pdf")

            with open(output_path, "wb") as output_pdf:
                pdf_writer.write(output_pdf)

@app.route('/download')
def download():
    # Create a zip file of the output folder
    shutil.make_archive('output_files', 'zip', app.config['OUTPUT_FOLDER'])
    return send_file('output_files.zip', as_attachment=True)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
