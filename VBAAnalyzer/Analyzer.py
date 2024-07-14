from flask import Flask, request, render_template
import subprocess

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/analyze', methods=['POST'])
def analyze():
    if 'file' not in request.files:
        return 'No file part'
    file = request.files['file']
    if file.filename == '':
        return 'No selected file'
    if file:
        file_path = 'uploaded_file.xlsm'
        file.save(file_path)
        result = subprocess.run(['python', 'extract_vba.py', file_path], capture_output=True, text=True)
        return render_template('result.html', result=result.stdout)

if __name__ == '__main__':
    app.run(debug=True)
