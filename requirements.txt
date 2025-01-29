import os
import subprocess
from flask import Flask, request, send_file
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

@app.route('/')
def home():
    return "PPT to PDF Converter API is running. Use /convert to upload files."

@app.route('/convert', methods=['POST'])
def convert_ppt_to_pdf():
    if 'file' not in request.files:
        return {'error': 'No file part'}, 400
    
    file = request.files['file']
    if file.filename == '':
        return {'error': 'No selected file'}, 400
    
    filename = secure_filename(file.filename)
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(file_path)
    
    output_pdf = os.path.join(app.config['OUTPUT_FOLDER'], filename.replace('.pptx', '.pdf'))
    
    try:
        command = f"libreoffice --headless --convert-to pdf --outdir {app.config['OUTPUT_FOLDER']} {file_path}"
        subprocess.run(command, shell=True, check=True)
        return send_file(output_pdf, as_attachment=True)
    except Exception as e:
        return {'error': str(e)}, 500

if __name__ == '__main__':
    app.run(debug=True)
