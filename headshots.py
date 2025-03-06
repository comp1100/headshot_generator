""" 
This takes an enrolment spreadsheet downloaded from Allocate+, and for each tutorial/workshop/studio, 
it creates a PDF file with the headshots and names of students enrolled. 
Steps to use:
1. Download an xls of your class from allocate. Select only the tutorials/workshops/studios/whatever, and then select 'Class List'.
2. Download the headshots for your couse from my.uq.edu.au -> 'My Courses' -> Select course -> 'In Person', and 'Student headshots'. Download the zip file
3. Run this script using:  python3 headshots.py <course classlist>.xls <headshots zip file>
"""

import os
import sys
import zipfile
import xlrd
from fpdf import FPDF
from PIL import Image
from io import BytesIO

# Constants for layout
IMAGES_PER_ROW = 4
ROWS_PER_PAGE = 4
IMAGES_PER_PAGE = IMAGES_PER_ROW * ROWS_PER_PAGE

IMAGE_WIDTH = 40  # Adjust image width
IMAGE_HEIGHT = 50  # Adjust image height
MARGIN_X = 10  # Left margin
MARGIN_Y = 10  # Top margin
SPACING_X = 10  # Horizontal spacing between images
SPACING_Y = 10  # Vertical spacing between images
TEXT_HEIGHT = 6  # Space for filename text


def get_student_numbers(sheet):
    """Read student numbers starting from row 8 in the first column of the sheet."""
    student_numbers = []
    for row_idx in range(7, sheet.nrows):  # Start from row 8 (index 7)
        student_number = sheet.cell_value(row_idx, 0)  # First column (A)
        if student_number:
            student_number = str(student_number).strip()  # Convert to string and strip whitespace
            if '.' in student_number:
                student_number = student_number.rstrip('0').rstrip('.')  # Remove trailing zeros and decimal point
            student_numbers.append(student_number)
    return student_numbers


def find_matching_images(student_numbers, zip_path):
    """Find images inside the ZIP file that contain student numbers in their filenames."""
    images = {}

    with zipfile.ZipFile(zip_path, "r") as z:
        zip_files = z.namelist()  # List all files in ZIP

        for num in student_numbers:
            matched_files = [
                f
                for f in zip_files
                if num in f and f.lower().endswith((".png", ".jpg", ".jpeg"))
            ]
            for matched_file in matched_files:
                images[matched_file] = z.read(
                    matched_file
                )  # Store file content as bytes

    return images


def remove_file_extension(filename):
    """Remove the file extension from the filename."""
    return os.path.splitext(filename)[0]


def create_pdf(image_data, output_pdf):
    """Generate a PDF file with images and filenames in a 4x4 grid per page."""
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=10)

    # Convert the image data dictionary to a list for easy iteration
    image_list = list(image_data.items())

    # Iterate through all images, adding them in groups of 16 per page
    for i in range(0, len(image_list), IMAGES_PER_PAGE):
        pdf.add_page()

        # Get the images for this page
        page_images = image_list[i : i + IMAGES_PER_PAGE]

        # Process each image on the current page
        for j, (filename, img_bytes) in enumerate(page_images):
            row = j // IMAGES_PER_ROW
            col = j % IMAGES_PER_ROW

            # Calculate position for each image in the grid
            x = MARGIN_X + col * (IMAGE_WIDTH + SPACING_X)
            y = MARGIN_Y + row * (IMAGE_HEIGHT + SPACING_Y + TEXT_HEIGHT * 2)

            try:
                # Open image from bytes data
                img = Image.open(BytesIO(img_bytes))
                img_path = f"temp_{i + j}.jpg"  # Ensure unique temporary filename based on both page index (i) and image index (j)
                img.save(img_path)  # Temporarily save to disk

                # Add the image to the PDF
                pdf.image(img_path, x=x, y=y, w=IMAGE_WIDTH, h=IMAGE_HEIGHT)

                # Add the filename as the label (remove the file extension)
                label = remove_file_extension(filename)
                name, student_number = label.rsplit("(", 1)  # Split by the last '('
                student_number = student_number.rstrip(")")  # Remove the trailing ')'

                # Set the x and y positions for the text
                pdf.set_xy(x, y + IMAGE_HEIGHT)
                pdf.set_font("Helvetica", size=8)
                pdf.cell(
                    IMAGE_WIDTH, TEXT_HEIGHT - 2, name.strip(), align="C"
                )  # Reduced height for name
                pdf.ln(TEXT_HEIGHT - 2)  # Reduced line height
                pdf.set_x(x)
                pdf.cell(
                    IMAGE_WIDTH, TEXT_HEIGHT - 2, student_number, align="C"
                )  # Reduced height for student number

                os.remove(img_path)  # Clean up temporary file

            except Exception as e:
                print(f"Error processing {filename}: {e}")

    if len(pdf.pages) > 0:
        pdf.output(output_pdf)
        print(f"PDF saved as {output_pdf}")
    else:
        print("No valid images found. No PDF created.")


def extract_pdf_name(cell_value):
    """Extract the PDF name from the cell value in the format 'Day_HH:MM'."""
    parts = cell_value.split(",")
    if len(parts) >= 2:
        day = parts[0].strip()
        time = parts[1].strip()
        return f"{day}_{time.replace(':', '')}"
    return "Unknown"


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python script.py student_list.xls image_archive.zip")
        sys.exit(1)

    xls_file = sys.argv[1]
    zip_file = sys.argv[2]

    if not os.path.exists(zip_file) or not zipfile.is_zipfile(zip_file):
        print(f"Error: The specified file '{zip_file}' is not a valid ZIP archive.")
        sys.exit(1)

    if not os.path.exists(xls_file):
        print(f"Error: The specified file '{xls_file}' does not exist.")
        sys.exit(1)

    # Load the .xls file using xlrd
    workbook = xlrd.open_workbook(xls_file)

    # Loop through each sheet in the workbook
    for sheet_name in workbook.sheet_names():
        sheet = workbook.sheet_by_name(sheet_name)

        # Get the student numbers from the sheet
        student_numbers = get_student_numbers(sheet)

        print(student_numbers) 
        # Get the name for the PDF from cell B3
        pdf_name = extract_pdf_name(sheet.cell_value(2, 1))  # Cell B3 is at (2, 1)
        if not pdf_name:
            print(
                f"Warning: No name found in cell B3 for sheet '{sheet_name}'. Skipping."
            )
            continue

        # Find matching images in the ZIP file
        images = find_matching_images(student_numbers, zip_file)

        # Generate the PDF for the sheet
        pdf_file = f"{pdf_name}.pdf"  # Name PDF based on cell B3 value
        if images:
            create_pdf(images, pdf_file)
        else:
            print(f"No matching images found for sheet '{sheet_name}'. No PDF created.")
