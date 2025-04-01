import os
import re
from datetime import datetime
from typing import Optional, List, Tuple, Dict, Any

import requests
import numpy as np
import gspread
from bs4 import BeautifulSoup
from pdf2image import convert_from_path
import easyocr
from PIL import Image

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

def save_image(img: Image.Image, page_num: int) -> str:
    """Save image to local storage.
    
    Args:
        img: PIL Image object
        page_num: Page number for filename
        keepImage: Whether to keep the image after processing
        
    Returns:
        Path to saved image
    """
    image_path = f"images/page_{page_num}.png"
    os.makedirs(os.path.dirname(image_path), exist_ok=True)
    img.save(image_path, "PNG")
    return image_path

def select_text_in_region(ocr_results: List[Tuple], region: Tuple[int, int, int, int], confidence_threshold: float = 0.5) -> List[Tuple[str, float]]:
    """Select text that falls within the specified region.
    
    Args:
        ocr_results: List of OCR results (box, text, confidence)
        region: Tuple of (x1, y1, x2, y2) coordinates
        confidence_threshold: Minimum confidence score for text selection
        
    Returns:
        List of tuples containing (text, confidence) for selected text
    """
    found_texts = []
    for detection in ocr_results:
        box = detection[0]  # Bounding box coordinates
        text = detection[1]  # The actual text content
        confidence = detection[2]  # Confidence score
        
        try:
            if is_box_in_region(box, region) and confidence > confidence_threshold:
                found_texts.append((text, confidence))
        except Exception as e:
            print(f"Warning: Error checking box {box}: {str(e)}")
            continue
    
    return found_texts

def get_page_range() -> Tuple[int, Optional[int]]:
    """Get start and end page numbers from user input."""
    while True:
        try:
            start_page = int(input("Enter start page number (1 or greater): "))
            if start_page < 1:
                print("Start page must be greater than 0")
                continue
                
            end_page_input = input("Enter end page number (press Enter for all remaining pages): ").strip()
            if not end_page_input:
                return start_page, None
                
            end_page = int(end_page_input)
            if end_page < start_page:
                print("End page must be greater than or equal to start page")
                continue
                
            return start_page, end_page
            
        except ValueError:
            print("Please enter valid numbers")

def process_images_only(pdf_path: str, start_page: int = 1, end_page: Optional[int] = None) -> None:
    """Process PDF and save specified pages as images without text extraction.
    
    Args:
        pdf_path: Path to the PDF file
        start_page: First page to process (1-based)
        end_page: Last page to process (1-based), None for all pages
    """
    try:
        # Convert PDF to images
        images = convert_from_path(pdf_path, poppler_path=POPPLER_PATH)
        total_pages = len(images)
        
        # Validate page numbers
        if start_page < 1:
            print("Error: Start page must be greater than 0")
            return
            
        if end_page is None:
            end_page = total_pages
        elif end_page > total_pages:
            print(f"Warning: End page {end_page} exceeds total pages {total_pages}. Using last page.")
            end_page = total_pages
        elif end_page < start_page:
            print("Error: End page must be greater than or equal to start page")
            return
            
        print(f"\nProcessing PDF pages {start_page} to {end_page} of {total_pages}...")
        
        # Save specified pages as images
        for i in range(start_page - 1, end_page):
            img = images[i]
            image_path = save_image(img, i + 1)
            print(f"Saved page {i + 1} to {image_path}")
            
        print(f"\nAll pages from {start_page} to {end_page} have been saved as images.")
        
    except Exception as e:
        print(f"Error processing PDF: {str(e)}")
        print("\nPlease make sure you have:")
        print("1. Installed Poppler from: https://github.com/oschwartz10612/poppler-windows/releases/")
        print("2. Set the correct POPPLER_PATH in the code to match your installation")

def process_text_only(pdf_path: str, configs: List[str]) -> List[Dict[str, str]]:
    """Process PDF and extract text without saving images."""
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
            
            # Convert PIL image to numpy array
            img_np = np.array(img)
            
            # Detect and recognize text once for all regions on this page
            ocr_results = reader.readtext(img_np)
            
            # Process each region on this page
            for config in configs:
                # Select text in the region
                found_texts = select_text_in_region(ocr_results, config['region'])
                
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
        
        return results
        
    except Exception as e:
        print(f"Error processing PDF: {str(e)}")
        print("\nPlease make sure you have:")
        print("1. Installed Poppler from: https://github.com/oschwartz10612/poppler-windows/releases/")
        print("2. Set the correct POPPLER_PATH in the code to match your installation")
        print("3. Installed EasyOCR using: pip install easyocr")
        return []

def process_company_data(choice: int = 1) -> None:
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
            
            if choice == 1:
                # Get page range from user
                start_page, end_page = get_page_range()
                # Process images only
                process_images_only(pdf_file_path, start_page, end_page)
            else:
                # Get all configurations
                configs = [config for config in row[2:] if config]  # Skip empty configs
                
                # Create or get the results worksheet
                try:
                    results_worksheet = sheet.worksheet(code)
                except:
                    results_worksheet = sheet.add_worksheet(code, 1000, 10)
                
                # Process text only
                results = process_text_only(pdf_file_path, configs)
                
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

def display_menu() -> None:
    """Display the main menu and get user choice."""
    while True:
        print("\nPDF Processing Menu")
        print("1. Save PDF pages as images")
        print("2. Extract text from PDF")
        print("3. Exit")
        
        try:
            choice = int(input("\nEnter your choice (1-3): "))
            if choice == 3:
                print("Goodbye!")
                break
            elif choice in [1, 2]:
                process_company_data(choice)
            else:
                print("Invalid choice. Please enter 1, 2, or 3.")
        except ValueError:
            print("Invalid input. Please enter a number.")

if __name__ == "__main__":
    display_menu()
