
# File Converter API

Convert between CSV and Excel formats with a simple web interface and REST API.

## Features

- Convert CSV to Excel (.xlsx)
- Convert Excel to CSV
- Web-based drag-and-drop interface
- REST API for programmatic access
- 16MB file size limit

## Quick Start

### Installation

```bash
pip install -r requirements.txt
```

### Run Locally

```bash
python app.py
```

Access the web interface at `http://localhost:10000`

## API Documentation

### Base URL
```
http://localhost:10000
```

### Endpoints

#### 1. Convert File
Convert between CSV and Excel formats.

**Endpoint:** `POST /convert`

**Request:**
- Method: `POST`
- Content-Type: `multipart/form-data`
- Body: Form data with file field

**Example (cURL):**
```bash
curl -X POST http://localhost:10000/convert \
  -F "file=@input.csv"
```

**Example (Python):**
```python
import requests

url = "http://localhost:10000/convert"
files = {"file": open("input.csv", "rb")}
response = requests.post(url, files=files)

with open("output.xlsx", "wb") as f:
    f.write(response.content)
```

**Example (JavaScript):**
```javascript
const formData = new FormData();
formData.append('file', fileInput.files[0]);

fetch('http://localhost:10000/convert', {
  method: 'POST',
  body: formData
})
.then(response => response.blob())
.then(blob => {
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'converted_file.xlsx';
  a.click();
});
```

**Response:**
- Success (200): File download with appropriate MIME type
- Error (400): `{"error": "error message"}`
- Error (500): `{"error": "conversion error message"}`

**Supported Formats:**
- Input: `.csv`, `.xlsx`, `.xls`
- Output: Automatically determined (CSV → Excel, Excel → CSV)

---

#### 2. Health Check
Check if the API is running.

**Endpoint:** `GET /health`

**Response:**
```json
{
  "status": "healthy"
}
```

---

#### 3. API Info
Get information about available endpoints.

**Endpoint:** `GET /api`

**Response:**
```json
{
  "message": "File Conversion API",
  "endpoints": {
    "/": "GET - Web interface",
    "/convert": "POST - Convert files",
    "/health": "GET - Health check"
  }
}
```

## File Requirements

- **Maximum file size:** 16MB
- **Supported input formats:** CSV, XLS, XLSX
- **Output formats:** XLSX (from CSV), CSV (from Excel)

## Error Handling

The API returns JSON error messages with appropriate HTTP status codes:

| Status Code | Description |
|-------------|-------------|
| 200 | Success - file converted |
| 400 | Bad request - invalid file or missing file |
| 500 | Server error - conversion failed |

**Example Error Response:**
```json
{
  "error": "Invalid file type. Allowed: csv, xlsx, xls"
}
```

## Deployment

### Deploy to Render

1. Push code to GitHub
2. Create new Web Service on Render
3. Connect repository
4. Deploy automatically with `render.yaml`

The service will be available at your Render URL.

## Tech Stack

- Flask - Web framework
- Pandas - Data processing
- OpenPyXL - Excel file handling
- Gunicorn - Production server

## Project Structure

```
file-converter/
├── app.py              # Main application
├── requirements.txt    # Python dependencies
├── render.yaml        # Render deployment config
└── README.md          # Documentation
```
