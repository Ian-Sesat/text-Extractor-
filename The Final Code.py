import os
import fitz
import pandas as pd
import numpy as np
from PIL import Image
from skimage.measure import find_contours
import re
import string

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
    dwg_number = '2021-019060'
    dwg_number_pattern = re.compile(r'dwg\.?\s*no\.?\s*(\d+-\d+)', re.IGNORECASE)
    text = page.get_text()
    match = dwg_number_pattern.search(text)
    if match:
        dwg_number = match.group(1)
    else:
        print(f"DWG number not found in page {page.number + 1}.")
    return dwg_number


def save_to_text_file(data, dwg_number):
    """
    Save extracted data to a text file based on DWG number.
    """
    output_file = f'{dwg_number}.txt'
    with open(output_file, 'w', encoding='utf-8') as file:
        for i, (box_data, _) in enumerate(data, start=1):
            # Remove illegal characters
            box_data = ''.join(filter(lambda x: x in string.printable, box_data))
            # Extract the box name
            box_name = re.search(r'(P\d+)', box_data)
            if box_name:
                box_name = box_name.group(1)
                box_data = box_data.replace(box_name, f"Box {box_name}")
            # Add colon after box name and before newline, with no extra space
            box_data = box_data.strip().replace('\n', '\n\n', 1).replace('\n', ': \n', 1)
            file.write(f"{box_data.strip()}\n\n")
    print(f"Extracted data saved to '{output_file}'.")

def save_to_excel(data, excel_file):
    """
    Save extracted data to an Excel file.
    """
    df = pd.DataFrame(data, columns=['Content', 'Bounding Box'])
    # Remove illegal characters
    df['Content'] = df['Content'].apply(lambda x: ''.join(filter(lambda y: y in string.printable, x)))
    df['Box Number'] = df.index + 1
    # Reorder columns
    df = df[['Box Number', 'Content', 'Bounding Box']]
    df.to_excel(excel_file, index=False)
    print(f"Combined data saved to '{excel_file}'.")

def main():
    # Replace 'pdf_directory' with the directory containing the PDF files
    pdf_directory = 'C:/Users/User/Desktop/Python Assignment'
    combined_data = []
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
                        save_to_text_file(black_boxes, dwg_number)
                        combined_data.extend(black_boxes)
                        print(f"Extracted data from page {page_num + 1} in '{pdf_file}'.")

    # Save combined data to an Excel file
    excel_file = 'combined_data.xlsx'
    save_to_excel(combined_data, excel_file)

if __name__ == "__main__":
    main()
