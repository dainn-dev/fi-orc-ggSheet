import os
import re
from datetime import datetime
from typing import Optional, List, Tuple, Dict, Any

import requests
import numpy as np
import gspread
from bs4 import BeautifulSoup
from PIL import Image
from pdf2image import convert_from_path
import easyocr

# Configuration
POPPLER_PATH = r"C:\Program Files\poppler-24.08.0\Library\bin"
SHEET_ID = "1O135ZG0rujXpki_zJyJxiVZhjWpDq2ytWc5Ra2hPyUE"
CREDENTIALS_FILE = os.path.join(os.path.dirname(__file__), "credentials.json")
WEB = "https://cafef.vn/du-lieu/Ajax/CongTy/BaoCaoTaiChinh.aspx?sym="

def clean_text(text: str) -> str:
    """Clean and normalize text by removing extra spaces and special characters."""
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'[^\w\s\u00C0-\u00FF.,;:()%]', '', text)
    text = re.sub(r'\n\s*\n', '\n\n', text)
    return text.strip()

def download_pdf(url: str, save_path: str) -> bool:
    """Download PDF from URL and save to specified path."""
    try:
        response = requests.get(url, stream=True)
        if response.status_code == 200:
            with open(save_path, "wb") as pdf_file:
                for chunk in response.iter_content(1024):
                    pdf_file.write(chunk)
            print(f"PDF downloaded successfully: {save_path}")
            return True
        print(f"Failed to download PDF. Status code: {response.status_code}")
        return False
    except Exception as e:
        print(f"Error downloading PDF: {str(e)}")
        return False

def is_box_in_region(box: List[Tuple[float, float]], region: Tuple[int, int, int, int]) -> bool:
    """Check if a text box intersects with the target region."""
    x1, y1, x2, y2 = region
    box = np.array(box, dtype=np.float32)
    
    # Check if any corner of the text box is inside the region
    for point in box:
        if (x1 <= point[0] <= x2) and (y1 <= point[1] <= y2):
            return True
    return False

def get_sheet_data() -> Optional[gspread.Spreadsheet]:
    """Initialize and return Google Sheets client."""
    try:
        # Check if credentials file exists
        if not os.path.exists(CREDENTIALS_FILE):
            print(f"Error: Credentials file not found at {CREDENTIALS_FILE}")
            print("Please make sure you have:")
            print("1. Downloaded the credentials.json file from Google Cloud Console")
            print("2. Placed it in the same directory as this script")
            print("3. Named it 'credentials.json'")
            return None

        # Validate credentials file format and print service account email
        try:
            import json
            with open(CREDENTIALS_FILE, 'r') as f:
                creds = json.load(f)
                required_fields = ['type', 'project_id', 'private_key', 'client_email', 'token_uri']
                missing_fields = [field for field in required_fields if field not in creds]
                if missing_fields:
                    print(f"Error: Credentials file is missing required fields: {', '.join(missing_fields)}")
                    print("Please make sure you downloaded the correct credentials file from Google Cloud Console")
                    return None
                print(f"\nService Account Email: {creds['client_email']}")
                print("Please share your Google Sheet with this email address")
        except json.JSONDecodeError:
            print("Error: Credentials file is not valid JSON")
            return None
        except Exception as e:
            print(f"Error reading credentials file: {str(e)}")
            return None

        # Initialize the Google Sheets client
        gc = gspread.service_account(filename=CREDENTIALS_FILE)
        
        # Try to open the sheet
        try:
            sheet = gc.open_by_key(SHEET_ID)
            print("Successfully connected to Google Sheet")
            return sheet
        except gspread.exceptions.SpreadsheetNotFound:
            print(f"Error: Could not find spreadsheet with ID {SHEET_ID}")
            print("Please make sure:")
            print("1. The SHEET_ID is correct")
            print("2. The service account has access to the spreadsheet")
            return None
        except gspread.exceptions.APIError as e:
            print(f"Error: API Error - {str(e)}")
            print("\nPlease make sure you have:")
            print("1. Shared the Google Sheet with the service account email shown above")
            print("2. Given the service account 'Editor' access")
            print("3. Enabled the Google Sheets API in your Google Cloud Console")
            return None
            
    except Exception as e:
        print(f"Error accessing Google Sheets: {str(e)}")
        print("\nPlease make sure you have:")
        print("1. Installed gspread using: pip install gspread")
        print("2. Set up Google Cloud Project and enabled Google Sheets API")
        print("3. Created a service account and downloaded credentials.json")
        return None

def get_pdf_file(code: str) -> Optional[str]:
    """Get PDF URL for a given company code."""
    try:
        response = requests.get(WEB + code)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            tables = soup.find_all('table')
            
            if len(tables) > 1:
                second_table = tables[1]
                rows = second_table.find_all('tr')
                
                for row in rows:
                    if "báo cáo tài chính hợp nhất" in row.text.lower():
                        print("Matching <tr> Found:", row)
                        link = row.find('a', href=True)
                        
                        if link and link['href'].endswith('.pdf'):
                            pdf_url = link['href']
                            if not pdf_url.startswith("http"):
                                pdf_url = "https://cafef.vn" + pdf_url
                            print("Extracted PDF URL:", pdf_url)
                            return pdf_url
                        else:
                            print("No PDF URL found in the matching row.")
                
            else:
                print("No table elements found.")
                return None
        else:
            print(f"Error fetching the URL: {response.status_code}")
            return None
    except Exception as e:
        print("Error fetching the URL:", e)
        return None

def extract_text_from_pdf(pdf_path: str, configs: List[str], keepImage: bool = False) -> List[Dict[str, str]]:
    """Extract text from PDF for multiple regions at once.
    
    Args:
        pdf_path: Path to the PDF file
        configs: List of configuration strings in format "page,target_cell,x1,y1,x2,y2"
        
    Returns:
        List of dictionaries containing target_cell and extracted text
    """
    try:
        # Initialize EasyOCR with Vietnamese and English support
        reader = easyocr.Reader(['vi', 'en'])
        
        # Convert PDF to images once for all regions
        images = convert_from_path(pdf_path, poppler_path=POPPLER_PATH)
        
        # Group configs by page number to minimize image processing
        page_configs = {}
        for config in configs:
            if not config:
                continue
            parts = config.split(',')
            if len(parts) < 6:
                continue
                
            page_num = int(parts[0])
            if page_num not in page_configs:
                page_configs[page_num] = []
            page_configs[page_num].append({
                'target_cell': parts[1],
                'region': (
                    int(parts[2]),  # x1
                    int(parts[3]),  # y1
                    int(parts[4]),  # x2
                    int(parts[5])   # y2
                )
            })
        
        results = []
        # Process each page once
        for page_num, configs in page_configs.items():
            # Adjust page number to 0-based index
            page_index = page_num - 1
            
            if page_index < 0 or page_index >= len(images):
                print(f"Error: Page {page_num} does not exist in the PDF (total pages: {len(images)})")
                continue
                
            # Get the target page image
            img = images[page_index]
            
            # Save the image locally
            image_path = f"images/page_{page_num}.png"
            os.makedirs(os.path.dirname(image_path), exist_ok=True)
            img.save(image_path, "PNG")
            
            # Convert PIL image to numpy array
            img_np = np.array(img)
            
            # Detect and recognize text once for all regions on this page
            ocr_results = reader.readtext(img_np)
            
            # Process each region on this page
            for config in configs:
                found_texts = []
                for detection in ocr_results:
                    box = detection[0]  # Bounding box coordinates
                    text = detection[1]  # The actual text content
                    confidence = detection[2]  # Confidence score
                    
                    try:
                        # Check if the text box intersects with the target region
                        if is_box_in_region(box, config['region']) and confidence > 0.5:
                            found_texts.append((text, confidence))
                    except Exception as e:
                        print(f"Warning: Error checking box {box}: {str(e)}")
                        continue
                
                if found_texts:
                    # Sort texts by y-coordinate to maintain reading order
                    found_texts.sort(key=lambda x: x[1])
                    combined_text = " ".join(text for text, _ in found_texts)
                    results.append({
                        'target_cell': config['target_cell'],
                        'text': combined_text
                    })
                else:
                    results.append({
                        'target_cell': config['target_cell'],
                        'text': "No text found in the specified region"
                    })
            
            # Clean up the image file
            if not keepImage:
                os.remove(image_path)
        
        return results
        
    except Exception as e:
        print(f"Error processing PDF: {str(e)}")
        print("\nPlease make sure you have:")
        print("1. Installed Poppler from: https://github.com/oschwartz10612/poppler-windows/releases/")
        print("2. Set the correct POPPLER_PATH in the code to match your installation")
        print("3. Installed EasyOCR using: pip install easyocr")
        return []

def process_company_data(keepImage: bool = False) -> None:
    """Process company data from Google Sheets and extract text from PDFs."""
    # Get the Google Sheet
    sheet = get_sheet_data()
    if not sheet:
        return
    
    # Get the company list worksheet
    company_worksheet = sheet.worksheet('Company')
    company_data = company_worksheet.get_all_values()
    
    # Process each company
    for row in company_data[1:]:  # Skip header row
        try:
            code = row[1]  # Company code is in second column
            print(f"\nProcessing company code: {code}")
            
            # Get PDF URL
            pdf_url = get_pdf_file(code)
            if not pdf_url:
                print(f"Could not find PDF URL for code {code}")
                continue
            
            # Download PDF
            pdf_file_path = f"reports/{pdf_url.split('/')[-1]}"
            os.makedirs(os.path.dirname(pdf_file_path), exist_ok=True)
            if not download_pdf(pdf_url, pdf_file_path):
                continue
            
            # Get all configurations
            configs = [config for config in row[2:] if config]  # Skip empty configs
            
            # Create or get the results worksheet
            try:
                results_worksheet = sheet.worksheet(code)
            except:
                results_worksheet = sheet.add_worksheet(code, 1000, 10)
            
            # Process all regions at once
            results = extract_text_from_pdf(pdf_file_path, configs, keepImage)
            
            # Save results to specific cells
            for result in results:
                try:
                    target_cell = result['target_cell']
                    text = result['text']
                    
                    # Update cell with text (one cell below)
                    next_row = int(target_cell[1:])
                    next_col = ord(target_cell[0].upper()) - ord('A') + 1
                    results_worksheet.update_cell(
                        next_row,
                        next_col,
                        text
                    )
                    
                    print(f"Updated cells {target_cell} and {chr(ord(target_cell[0]))}{int(target_cell[1:]) + 1} with results")
                except Exception as e:
                    print(f"Error updating cells: {str(e)}")
            
            # Clean up
            os.remove(pdf_file_path)
            print(f"Completed processing company {code}")

        except Exception as e:
            print(f"Error processing company {code}: {str(e)}")
            continue

if __name__ == "__main__":
    process_company_data(keepImage=True)
