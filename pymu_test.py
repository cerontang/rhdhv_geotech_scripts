import fitz  # PyMuPDF
import cv2
import numpy as np


# Function to find the outermost black boundary in an image
def find_outer_boundary(image):
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(gray, 1, 255, cv2.THRESH_BINARY)
    contours, _ = cv2.findContours(binary, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    if len(contours) == 0:
        return None

    max_contour = max(contours, key=cv2.contourArea)
    x, y, w, h = cv2.boundingRect(max_contour)

    return x, y, x + w, y + h


# Specify the input and output file paths
input_pdf = r"C:\Users\921722\OneDrive - Royal HaskoningDHV\20230928 Extraction\PSD\Sieve\test\PSD_Sieve.pdf"
output_pdf = r"C:\Users\921722\OneDrive - Royal HaskoningDHV\20230928 Extraction\PSD\Sieve\test\PSD_Sieve_cropped.pdf"

# Open the input PDF file you want to process
pdf_document = fitz.open(input_pdf)

# Create a PDF writer object to write the cropped pages to a new PDF
pdf_writer = fitz.open()

# Loop through each page in the PDF
for page_number in range(len(pdf_document)):
    # Get the page
    page = pdf_document[page_number]

    # Convert the PDF page to an image
    pix = page.get_pixmap()
    img_data = pix.samples
    img = cv2.imdecode(np.frombuffer(img_data, dtype=np.uint8), -1)

    # Find the outermost black boundary
    boundary = find_outer_boundary(img)

    if boundary:
        left, top, right, bottom = boundary

        # Crop the image to the boundary
        img_cropped = img[top:bottom, left:right]

        # Create a new PDF page and add the cropped image
        new_page = pdf_writer.new_page(width=img_cropped.shape[1], height=img_cropped.shape[0])
        new_page.set_pixmap(fitz.Pixmap(img_cropped))

# Save the cropped pages to a new PDF file
pdf_writer.save(output_pdf)
pdf_writer.close()

print(f"Cropped PDF saved to {output_pdf}")
