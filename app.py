from flask import Flask, request, send_file, jsonify, render_template_string
from flask_cors import CORS
import pandas as pd
import io
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
CORS(app)

app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

ALLOWED_EXTENSIONS = {'csv', 'xlsx', 'xls'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# HTML Template
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>File Converter - CSV ↔ Excel</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            justify-content: center;
            align-items: center;
            padding: 20px;
        }

        .container {
            background: white;
            border-radius: 20px;
            padding: 40px;
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3);
            max-width: 500px;
            width: 100%;
        }

        h1 {
            text-align: center;
            color: #333;
            margin-bottom: 10px;
            font-size: 28px;
        }

        .subtitle {
            text-align: center;
            color: #666;
            margin-bottom: 30px;
            font-size: 14px;
        }

        .upload-area {
            border: 3px dashed #667eea;
            border-radius: 15px;
            padding: 40px 20px;
            text-align: center;
            cursor: pointer;
            transition: all 0.3s ease;
            background: #f8f9ff;
        }

        .upload-area:hover {
            border-color: #764ba2;
            background: #f0f2ff;
        }

        .upload-area.dragging {
            border-color: #764ba2;
            background: #e8ebff;
            transform: scale(1.02);
        }

        .upload-icon {
            font-size: 48px;
            margin-bottom: 15px;
        }

        .upload-text {
            color: #667eea;
            font-weight: 600;
            margin-bottom: 5px;
        }

        .upload-subtext {
            color: #999;
            font-size: 13px;
        }

        input[type="file"] {
            display: none;
        }

        .file-info {
            margin-top: 20px;
            padding: 15px;
            background: #f0f2ff;
            border-radius: 10px;
            display: none;
        }

        .file-info.active {
            display: block;
        }

        .file-name {
            color: #333;
            font-weight: 600;
            margin-bottom: 5px;
            word-break: break-all;
        }

        .file-size {
            color: #666;
            font-size: 13px;
        }

        .convert-btn {
            width: 100%;
            padding: 15px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: none;
            border-radius: 10px;
            font-size: 16px;
            font-weight: 600;
            cursor: pointer;
            margin-top: 20px;
            transition: transform 0.2s ease;
            display: none;
        }

        .convert-btn.active {
            display: block;
        }

        .convert-btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(102, 126, 234, 0.4);
        }

        .convert-btn:disabled {
            background: #ccc;
            cursor: not-allowed;
            transform: none;
        }

        .status {
            margin-top: 20px;
            padding: 15px;
            border-radius: 10px;
            text-align: center;
            display: none;
        }

        .status.active {
            display: block;
        }

        .status.success {
            background: #d4edda;
            color: #155724;
        }

        .status.error {
            background: #f8d7da;
            color: #721c24;
        }

        .status.loading {
            background: #fff3cd;
            color: #856404;
        }

        .spinner {
            border: 3px solid #f3f3f3;
            border-top: 3px solid #667eea;
            border-radius: 50%;
            width: 30px;
            height: 30px;
            animation: spin 1s linear infinite;
            margin: 0 auto 10px;
        }

        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }

        .supported-formats {
            text-align: center;
            margin-top: 20px;
            padding-top: 20px;
            border-top: 1px solid #eee;
        }

        .supported-formats h3 {
            font-size: 14px;
            color: #666;
            margin-bottom: 10px;
        }

        .format-badges {
            display: flex;
            justify-content: center;
            gap: 10px;
            flex-wrap: wrap;
        }

        .badge {
            background: #667eea;
            color: white;
            padding: 5px 15px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>📊 File Converter</h1>
        <p class="subtitle">Convert between CSV and Excel formats instantly</p>

        <div class="upload-area" id="uploadArea">
            <div class="upload-icon">📁</div>
            <div class="upload-text">Click to upload or drag and drop</div>
            <div class="upload-subtext">CSV, XLS, or XLSX files (Max 16MB)</div>
            <input type="file" id="fileInput" accept=".csv,.xlsx,.xls">
        </div>

        <div class="file-info" id="fileInfo">
            <div class="file-name" id="fileName"></div>
            <div class="file-size" id="fileSize"></div>
        </div>

        <button class="convert-btn" id="convertBtn">Convert File</button>

        <div class="status" id="status"></div>

        <div class="supported-formats">
            <h3>Supported Conversions</h3>
            <div class="format-badges">
                <span class="badge">CSV → Excel</span>
                <span class="badge">Excel → CSV</span>
            </div>
        </div>
    </div>

    <script>
        // API URL is automatically set to same domain
        const API_URL = window.location.origin;

        const uploadArea = document.getElementById('uploadArea');
        const fileInput = document.getElementById('fileInput');
        const fileInfo = document.getElementById('fileInfo');
        const fileName = document.getElementById('fileName');
        const fileSize = document.getElementById('fileSize');
        const convertBtn = document.getElementById('convertBtn');
        const status = document.getElementById('status');

        let selectedFile = null;

        uploadArea.addEventListener('click', () => fileInput.click());

        fileInput.addEventListener('change', (e) => {
            handleFile(e.target.files[0]);
        });

        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.classList.add('dragging');
        });

        uploadArea.addEventListener('dragleave', () => {
            uploadArea.classList.remove('dragging');
        });

        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('dragging');
            handleFile(e.dataTransfer.files[0]);
        });

        function handleFile(file) {
            if (!file) return;

            const validExtensions = ['.csv', '.xls', '.xlsx'];
            const fileExt = '.' + file.name.split('.').pop().toLowerCase();

            if (!validExtensions.includes(fileExt)) {
                showStatus('Please select a valid CSV or Excel file', 'error');
                return;
            }

            if (file.size > 16 * 1024 * 1024) {
                showStatus('File size must be less than 16MB', 'error');
                return;
            }

            selectedFile = file;
            fileName.textContent = file.name;
            fileSize.textContent = formatFileSize(file.size);
            fileInfo.classList.add('active');
            convertBtn.classList.add('active');
            hideStatus();
        }

        convertBtn.addEventListener('click', async () => {
            if (!selectedFile) return;

            const formData = new FormData();
            formData.append('file', selectedFile);

            convertBtn.disabled = true;
            showStatus('Converting your file...', 'loading', true);

            try {
                const response = await fetch(`${API_URL}/convert`, {
                    method: 'POST',
                    body: formData
                });

                if (!response.ok) {
                    const error = await response.json();
                    throw new Error(error.error || 'Conversion failed');
                }

                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                
                const contentDisposition = response.headers.get('Content-Disposition');
                let downloadFileName = 'converted_file';
                if (contentDisposition) {
                    const match = contentDisposition.match(/filename="?(.+)"?/);
                    if (match) downloadFileName = match[1];
                } else {
                    const originalExt = selectedFile.name.split('.').pop().toLowerCase();
                    const newExt = originalExt === 'csv' ? 'xlsx' : 'csv';
                    downloadFileName = selectedFile.name.replace(/\\.[^/.]+$/, '') + '.' + newExt;
                }
                
                a.download = downloadFileName;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);

                showStatus('✓ File converted successfully!', 'success');
                setTimeout(resetForm, 3000);
            } catch (error) {
                showStatus(`✗ Error: ${error.message}`, 'error');
            } finally {
                convertBtn.disabled = false;
            }
        });

        function showStatus(message, type, showSpinner = false) {
            status.className = `status active ${type}`;
            status.innerHTML = showSpinner ? 
                `<div class="spinner"></div>${message}` : 
                message;
        }

        function hideStatus() {
            status.classList.remove('active');
        }

        function formatFileSize(bytes) {
            if (bytes === 0) return '0 Bytes';
            const k = 1024;
            const sizes = ['Bytes', 'KB', 'MB'];
            const i = Math.floor(Math.log(bytes) / Math.log(k));
            return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
        }

        function resetForm() {
            selectedFile = null;
            fileInput.value = '';
            fileInfo.classList.remove('active');
            convertBtn.classList.remove('active');
            hideStatus();
        }
    </script>
</body>
</html>
'''

@app.route('/', methods=['GET'])
def home():
    return render_template_string(HTML_TEMPLATE)

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'healthy'})

@app.route('/api', methods=['GET'])
def api_info():
    return jsonify({
        'message': 'File Conversion API',
        'endpoints': {
            '/': 'GET - Web interface',
            '/convert': 'POST - Convert files (csv to excel or excel to csv)',
            '/health': 'GET - Health check'
        }
    })

@app.route('/convert', methods=['POST'])
def convert_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400

    if not allowed_file(file.filename):
        return jsonify({'error': 'Invalid file type. Allowed: csv, xlsx, xls'}), 400

    try:
        filename = secure_filename(file.filename)
        file_ext = filename.rsplit('.', 1)[1].lower()

        # Read the file based on its type
        if file_ext == 'csv':
            df = pd.read_csv(file)
            output_ext = 'xlsx'
            output_filename = filename.rsplit('.', 1)[0] + '.xlsx'

            # Convert to Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')
            output.seek(0)
            mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

        else:  # xlsx or xls
            df = pd.read_excel(file)
            output_ext = 'csv'
            output_filename = filename.rsplit('.', 1)[0] + '.csv'

            # Convert to CSV
            output = io.StringIO()
            df.to_csv(output, index=False)
            output.seek(0)
            output = io.BytesIO(output.getvalue().encode('utf-8'))
            mimetype = 'text/csv'

        return send_file(
            output,
            mimetype=mimetype,
            as_attachment=True,
            download_name=output_filename
        )

    except Exception as e:
        return jsonify({'error': f'Conversion failed: {str(e)}'}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)