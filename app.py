import os
from flask import Flask, request, redirect, render_template, send_file
from werkzeug.utils import secure_filename
from docx import Document
from reportlab.pdfgen import canvas
from pdf2docx import Converter

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
CONVERTED_FOLDER = 'converted'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CONVERTED_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    file = request.files['file']
    direction = request.form['direction']

    if file:
        filename = secure_filename(file.filename)
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        file.save(filepath)

        if direction == 'docx_to_pdf' and filename.endswith('.docx'):
            pdf_path = os.path.join(CONVERTED_FOLDER, filename.replace('.docx', '.pdf'))
            doc = Document(filepath)
            pdf = canvas.Canvas(pdf_path)
            y = 800
            for para in doc.paragraphs:
                pdf.drawString(100, y, para.text)
                y -= 20
            pdf.save()
            return send_file(pdf_path, as_attachment=True)

        elif direction == 'pdf_to_docx' and filename.endswith('.pdf'):
            docx_path = os.path.join(CONVERTED_FOLDER, filename.replace('.pdf', '.docx'))
            cv = Converter(filepath)
            cv.convert(docx_path, start=0, end=None)
            cv.close()
            return send_file(docx_path, as_attachment=True)
        
        else:
            return "Invalid file type or conversion direction."
    return redirect('/')

if __name__ == '__main__':
    app.run(debug=True)