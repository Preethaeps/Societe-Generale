from flask import Flask, request, jsonify, render_template
from werkzeug.utils import secure_filename
import os
import openai
from oletools.olevba import VBA_Parser

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ALLOWED_EXTENSIONS'] = {'xlsm', 'xlsx'}

openai.api_key = 'your-key'

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def extract_vba_macros(file_path):
    try:
        if not os.path.exists(file_path):
            return "File not found"

        vba_parser = VBA_Parser(file_path)
        if vba_parser.detect_vba_macros():
            macros = []
            for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_macros():
                macros.append({
                    "filename": filename,
                    "stream_path": stream_path,
                    "vba_filename": vba_filename,
                    "vba_code": vba_code
                })
            return macros
        else:
            return "No VBA macros found."
    except Exception as e:
        return str(e)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)

        macros = extract_vba_macros(file_path)
        return jsonify(macros)
    else:
        return jsonify({"error": "Invalid file format. Please upload a .xlsm or .xlsx file."}), 400

if __name__ == '__main__':
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    app.run(debug=True)
