from flask import Flask, render_template, request, send_file
import pdfplumber
import pandas as pd
from docx import Document
import fitz  # PyMuPDF
import io
import os

app = Flask(__name__, template_folder='../templates')

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return "No file uploaded", 400
    
    file = request.files['file']
    format_type = request.form.get('format')
    pdf_bytes = file.read()

    try:
        # --- IMAGE CONVERSION (PNG) ---
        if format_type == 'img':
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            page = doc.load_page(0)  # Convert first page
            pix = page.get_pixmap()
            img_data = pix.tobytes("png")
            
            return send_file(
                io.BytesIO(img_data),
                mimetype='image/png',
                as_attachment=True,
                download_name='converted_page.png'
            )

        # --- EXCEL CONVERSION (XLSX) ---
        elif format_type == 'xlsx':
            output = io.BytesIO()
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                all_data = []
                for page in pdf.pages:
                    table = page.extract_table()
                    if table:
                        all_data.extend(table)
                    else:
                        text = page.extract_text()
                        if text:
                            all_data.append([text])
            
            df = pd.DataFrame(all_data)
            df.to_excel(output, index=False, header=False)
            output.seek(0)
            return send_file(
                output,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name='converted_data.xlsx'
            )

        # --- WORD CONVERSION (DOCX) ---
        elif format_type == 'docx':
            output = io.BytesIO()
            word_doc = Document()
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        word_doc.add_paragraph(text)
            
            word_doc.save(output)
            output.seek(0)
            return send_file(
                output,
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                as_attachment=True,
                download_name='converted_text.docx'
            )

    except Exception as e:
        return f"Error: {str(e)}", 500

# Required for Vercel
app.debug = False