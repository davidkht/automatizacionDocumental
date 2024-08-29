from fpdf import FPDF
from PyPDF2 import PdfReader, PdfWriter
import os
from PIL import Image

class PDF(FPDF):
    def header(self):
        if self.page_no() == 1:
            self.set_font('Arial', 'B', 16)
            self.cell(0, 10, 'Registro Fotográfico', 0, 1, 'C')
            self.ln(10)

def reduce_image_quality(image_path, quality):
    with Image.open(image_path) as img:
        img = img.convert("RGB")  # Ensure compatibility
        temp_image_path = image_path + "_temp.jpg"
        img.save(temp_image_path, format='JPEG', quality=quality)
    return temp_image_path

def insert_images_to_pdf(image_folder, output_pdf):

    pdf = PDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    figure_number = 1
    images_per_row = 2
    images_per_page = 4
    images_count = 0
    quality = 30  # 95 - 20% = 76

    temp_files = []

    try:
        for image_file in os.listdir(image_folder):
            if image_file.lower().endswith(".jpg") or image_file.lower().endswith(".jpeg") or image_file.lower().endswith(".png"):
                if images_count % images_per_page == 0:
                    pdf.add_page()

                x_offset = 10 + (images_count % images_per_row) * 100
                y_offset = 30 + (images_count // images_per_row % (images_per_page // images_per_row)) * 140

                image_path = os.path.join(image_folder, image_file)
                temp_image_path = reduce_image_quality(image_path, quality)
                temp_files.append(temp_image_path)

                pdf.image(temp_image_path, x=x_offset, y=y_offset, w=90)

                pdf.set_xy(x_offset, y_offset + 90)  # Position just below the image
                pdf.set_font('Arial', '', 12)
                pdf.cell(90, 3, f'Figura {figure_number}', 0, 0, 'C')  # Reduced cell height to 3

                figure_number += 1
                images_count += 1
    finally:
        # Cleanup temporary files
        for temp_file in temp_files:
            os.remove(temp_file)
    
    pdf.output(output_pdf)

def merge_pdfs(pdf_list, output_pdf):
    pdf_writer = PdfWriter()

    for pdf in pdf_list:
        pdf_reader = PdfReader(pdf)
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            pdf_writer.add_page(page)
    
    with open(output_pdf, 'wb') as out:
        pdf_writer.write(out)