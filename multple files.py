import os
import fitz
import numpy as np
from PIL import Image
from skimage.measure import find_contours
import re

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

def save_to_text_file(data, output_file):
    """
    Save extracted data to a text file with sequential labeling.
    """
    with open(output_file, 'w', encoding='utf-8') as file:
        for i, (box_data, _) in enumerate(data, start=1):
            file.write(f"Box {i}:\n{box_data}\n\n")

def main():
    # Replace 'pdf_directory' with the directory containing the PDF files
    pdf_directory = 'C:/Users/User/Documents/PROGRAMMING/Python projects'
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
                        output_file = f'{dwg_number}.txt'
                        save_to_text_file(black_boxes, output_file)
                        print(f"Extracted data saved to '{output_file}' for page {page_num + 1} in '{pdf_file}'.")

if __name__ == "__main__":
    main()
