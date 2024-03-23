import os
import fitz
import numpy as np
from PIL import Image
from skimage.measure import find_contours
import re
import pandas as pd

def detect_black_boxes(page):
    """
    Detect black boxes within the PDF page.
    """
    black_boxes = []
    # Convert the page to an image and check for black regions
    image = page.get_pixmap()
    black_regions = find_black_regions(image)
    # Extract text from black regions
    for bbox in black_regions:
        text = page.get_text("text", clip=bbox)
        if text.strip().startswith('P') and len(text.strip()) > 20 and not text.strip().startswith('POLE'):
            matches = re.findall(r'\bP\d+\b', text)
            if matches:
                lines = text.strip().split('\n')
                filtered_lines = [line for line in lines if not any(word in line for word in ['DETAILS', 'NUMBER', 'NAME'])]
                filtered_text = '\n'.join(filtered_lines)
                black_boxes.append((filtered_text, bbox))
    return black_boxes

def find_black_regions(image):
    """
    Find black regions within an image.
    """
    # Convert the image to grayscale using Pillow
    pil_image = Image.frombytes("RGB", [image.width, image.height], image.samples)
    gray_image = pil_image.convert("L")
    # Threshold the image to obtain binary image
    binary_image = np.array(gray_image.point(lambda x: 0 if x < 200 else 255, '1'))
    # Get contours
    contours = find_contours(binary_image, 0.5)
    # Get bounding boxes of contours
    bounding_boxes = []
    for contour in contours:
        min_row, min_col = contour.min(axis=0)
        max_row, max_col = contour.max(axis=0)
        bounding_boxes.append((min_col, min_row, max_col, max_row))
    return bounding_boxes

def extract_dwg_number(page):
    """
    Extract the DWG number located at the bottom right of the PDF page.
    """
    dwg_number = None
    dwg_number_pattern = re.compile(r'dwg\.?\s*no\.?\s*(\d+-\d+)', re.IGNORECASE)
    text = page.get_text()
    match = dwg_number_pattern.search(text)
    if match:
        dwg_number = match.group(1)
    return dwg_number

def main():
    # Replace 'pdf_directory' with the directory containing the PDF files
    pdf_directory = 'C:/Users/User/Documents/PROGRAMMING/Python projects'
    all_data = []
    for pdf_file in os.listdir(pdf_directory):
        pdf_path = os.path.join(pdf_directory, pdf_file)
        if pdf_file.endswith('.pdf'):
            with fitz.open(pdf_path) as pdf_file:
                for page_num in range(len(pdf_file)):
                    page = pdf_file.load_page(page_num)
                    # Detect black boxes within the PDF page
                    black_boxes = detect_black_boxes(page)
                    if black_boxes:
                        dwg_number = extract_dwg_number(page)
                        for text, _ in black_boxes:
                            # Remove illegal characters
                            text = remove_illegal_characters(text)
                            all_data.append((dwg_number, text))

    # Convert extracted data to a Pandas DataFrame
    df = pd.DataFrame(all_data, columns=['DWG Number', 'Text'])

    # Save DataFrame to an Excel file
    excel_file = 'extracted_data.xlsx'
    df.to_excel(excel_file, index=False)
    print(f"Extracted data saved to '{excel_file}'.")

def remove_illegal_characters(text):
    """
    Remove illegal characters from text.
    """
    legal_chars = []
    for char in text:
        if char.isalnum() or char in [' ', ',', '.', '-', '_', '/', ':', ';', '(', ')']:
            legal_chars.append(char)
    return ''.join(legal_chars)



if __name__ == "__main__":
    main()
