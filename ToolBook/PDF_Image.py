import os
import fitz  # PyMuPDF
from PIL import Image

def pdf_to_images(pdf_path, output_folder, dpi=300):
    # Verifica si el directorio de salida existe, si no, créalo
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Abre el PDF
    pdf_document = fitz.open(pdf_path)
    
    # Itera sobre cada página del PDF
    for page_number in range(pdf_document.page_count):
        # Obtiene la página actual
        page = pdf_document.load_page(page_number)
        
        # Renderiza la página como una imagen con una resolución específica (dpi)
        image = page.get_pixmap(matrix=fitz.Matrix(dpi/72, dpi/72))
        
        # Convierte la imagen en formato RGB
        img = Image.frombytes("RGB", [image.width, image.height], image.samples)
        
        # Guarda la imagen en el directorio de salida con alta calidad (formato PNG)
        img.save(f"{output_folder}/page_{page_number + 1}.jpg", dpi=(dpi, dpi))

    # Cierra el documento PDF
    pdf_document.close()

# Ruta del archivo PDF de entrada (ruta completa)
input_pdf = "G:\\Shared drives\\ES VIALTO GMS - RPA\\TAX\\COMPLIANCE\\i_129s\\Templates\\i-129s.pdf"

# Carpeta de salida para las imágenes
output_folder = "G:\\Shared drives\\ES VIALTO GMS - RPA\\TAX\\COMPLIANCE\\i_129s\\Templates"

# Llama a la función para convertir el PDF en imágenes con una resolución de 600 dpi
pdf_to_images(input_pdf, output_folder, dpi=600)
