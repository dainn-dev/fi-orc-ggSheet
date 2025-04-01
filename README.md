# PDF Text Extractor with Google Sheets Integration

This Python script extracts text from specific regions of PDF files and saves the results to Google Sheets. It's designed to work with financial reports from cafef.vn and supports both Vietnamese and English text extraction.

## Prerequisites

### 1. Python Environment
- Python 3.7 or higher
- pip (Python package installer)

### 2. Required Python Packages
Install the following packages using pip:
```bash
pip install requests numpy opencv-python gspread beautifulsoup4 Pillow pdf2image easyocr
```

### 3. Poppler Installation
1. Download Poppler for Windows from: https://github.com/oschwartz10612/poppler-windows/releases/
2. Extract the downloaded file
3. Add the `bin` directory to your system's PATH environment variable
   - Default path: `C:\Program Files\poppler-24.08.0\Library\bin`

### 4. Google Cloud Setup
1. Create a Google Cloud Project
2. Enable the Google Sheets API
3. Create a Service Account:
   - Go to "IAM & Admin" > "Service Accounts"
   - Click "Create Service Account"
   - Give it a name and description
   - Grant "Editor" role
   - Create and download the JSON key file
   - Rename the downloaded file to `credentials.json`
   - Place it in the same directory as the script

### 5. Google Sheets Setup
1. Create a new Google Sheet
2. Share the sheet with the service account email (found in credentials.json)
3. Create a worksheet named "Company" with the following structure:
   - Column 1: Header row
   - Column 2: Company codes
   - Column 3+: Configuration strings in format "page,target_cell,x1,y1,x2,y2"
     - page: Page number in PDF (1-based)
     - target_cell: Cell reference (e.g., "B4")
     - x1,y1,x2,y2: Region coordinates(upload image to https://www.image-map.net/ and get coords)

## Configuration

### 1. Update Constants
In `main.py`, update the following constants:
```python
POPPLER_PATH = r"C:\Program Files\poppler-24.08.0\Library\bin"  # Update if different
SHEET_ID = "your_sheet_id_here"  # Your Google Sheet ID
CREDENTIALS_FILE = "credentials.json"  # Path to your credentials file
WEB = "https://cafef.vn/du-lieu/Ajax/CongTy/BaoCaoTaiChinh.aspx?sym="
```

### 2. Google Sheet Structure
Example of the "Company" worksheet:
```
| Header | Company Code | Config 1             | Config 2             |
|--------|------------- |----------------------|----------------------|
| Header | VNM          | 1,B4,100,100,200,200 | 2,B5,300,300,400,400 |
```

## Usage
0. Create folder images and reports
1. Run the script:
```bash
py main.py
```

2. The script will:
   - Read company codes from the "Company" worksheet
   - Download PDFs from cafef.vn
   - Extract text from specified regions
   - Save results to company-specific worksheets
   - Clean up temporary files

3. Results will be saved in:
   - Each company gets its own worksheet named after its code
   - Text is saved in the specified target cells
   - Images are saved in the `images` directory (if keepImage=True)

## Error Handling

The script includes error handling for:
- Missing credentials file
- Invalid Google Sheet access
- PDF download failures
- OCR processing errors
- Cell update errors

Check the console output for detailed error messages and troubleshooting steps.

## Notes

- The script requires internet connection to access cafef.vn and Google Sheets
- PDF processing may take some time depending on the number of pages and regions
- Make sure you have sufficient disk space for temporary files
- The script automatically cleans up temporary files after processing
