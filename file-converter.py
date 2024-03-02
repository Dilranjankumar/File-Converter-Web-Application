from flask import Flask, render_template, request, send_file
from pdf2docx import Converter
from docx2pdf import convert as docx_to_pdf_convert
from img2pdf import convert as img_to_pdf_convert
from PyPDF2 import PdfFileMerger
import os
import img2pdf

app = Flask(__name__)

# Function to check if the file extension is allowed
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'pdf', 'docx', 'jpg', 'jpeg', 'png', 'xlsx'}

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert_file():
    if 'file' not in request.files and 'files[]' not in request.files:
        return 'No file part'

    if 'file' in request.files:
        file = request.files['file']

        if file.filename == '':
            return 'No selected file'

        if file and allowed_file(file.filename):
            if file.filename.lower().endswith('.pdf'):
                # Convert PDF to Word
                pdf_path = os.path.join('F:\\Python Projects\\downloads', file.filename)
                docx_filename = file.filename.replace('.pdf', '.docx')
                docx_path = os.path.join('F:\\Python Projects\\downloads', docx_filename)
                os.makedirs(os.path.dirname(pdf_path), exist_ok=True)  # Create downloads directory if it doesn't exist
                file.save(pdf_path)
                
                with open(pdf_path, 'rb') as pdf_file:
                    with open(docx_path, 'wb') as docx_file:
                        cv = Converter(pdf_file)
                        cv.convert(docx_file, start=0, end=None)
                        cv.close()

                return send_file(docx_path, as_attachment=True)
            
            elif file.filename.lower().endswith('.docx'):
                # Convert Word to PDF
                docx_path = os.path.join('F:\\Python Projects\\downloads', file.filename)
                pdf_filename = file.filename.replace('.docx', '.pdf')
                pdf_path = os.path.join('F:\\Python Projects\\downloads', pdf_filename)
                os.makedirs(os.path.dirname(docx_path), exist_ok=True)  # Create downloads directory if it doesn't exist
                file.save(docx_path)
                docx_to_pdf_convert(docx_path, pdf_path)
                return send_file(pdf_path, as_attachment=True)

            elif file.filename.lower().endswith('.pdf'):
                # Convert PDF to Excel
                pdf_path = os.path.join('F:\\Python Projects\\downloads', file.filename)
                xlsx_filename = file.filename.replace('.pdf', '.xlsx')
                xlsx_path = os.path.join('F:\\Python Projects\\downloads', xlsx_filename)
                os.makedirs(os.path.dirname(pdf_path), exist_ok=True)  # Create downloads directory if it doesn't exist
                file.save(pdf_path)
                df = tabula.read_pdf(pdf_path, pages='all')
                df.to_excel(xlsx_path, index=False)
                return send_file(xlsx_path, as_attachment=True)

    if 'files[]' in request.files:
        files = request.files.getlist('files[]')
        pdfs = []
        for img_file in files:
            if img_file and allowed_file(img_file.filename):
                img_path = os.path.join('F:\\Python Projects\\downloads', img_file.filename)
                pdf_path = os.path.join('F:\\Python Projects\\downloads', img_file.filename.replace(img_file.filename.split('.')[-1], 'pdf'))
                img_file.save(img_path)
                with open(pdf_path, 'wb') as f:
                    f.write(img2pdf.convert(img_path))
                pdfs.append(pdf_path)

    if pdfs:
        pdf_filename = 'merged_images.pdf'
        pdf_path = os.path.join('F:\\Python Projects\\downloads', pdf_filename)
        merger = PdfFileMerger()

        for pdf in pdfs:
            merger.append(pdf)

        merger.write(pdf_path)
        merger.close()

        return send_file(pdf_path, as_attachment=True)

    return 'Invalid file format'

if __name__ == '__main__':
    app.run(debug=True)
