import os
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import pandas as pd
import openpyxl
from reportlab.lib.utils import ImageReader
import sys
import fitz  # PyMuPDF
from PIL import Image  # Para obtener dimensiones de la firma

# Añade el directorio PythonAutomation al sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from ToolBook import utils

# Rutas base para archivos y carpetas
ruta_base_user = 'G:/Shared drives/ES VIALTO GMS - RPA/INMI & SS/Model G28i/'
ruta_base_template = 'G:/Shared drives/ES VIALTO GMS - RPA/PythonAutomation/Bot_G028i/Templates/'
ruta_base = 'G:/Shared drives/ES VIALTO GMS - RPA/INMI & SS/FORM i129S/BOT - DO NOT TOUCH/'
output_folder = ruta_base_user + 'pdfs_generados'

def filter_all_na_columns(df):
    """Remove all-NA columns from a DataFrame."""
    return df.dropna(axis=1, how='all')

def g28_generate_pdfs_from_excel(excel_file, attorney_file, image_paths, output_folder):
    """
    Genera PDFs G28-i a partir de datos del archivo de excel del I129-s y attorney_info.
    """
    # Leer el archivo Excel
    xls = pd.ExcelFile(excel_file)
    data_temp = pd.read_excel(xls, 'Temp')
    data_log = pd.read_excel(xls, 'Log')
    
    # Leer el archivo attorney_info
    attorney_info = pd.read_excel(attorney_file)
    
    # Reemplazar NaN por cadenas vacías
    data_temp = data_temp.fillna('')
    attorney_info = attorney_info.fillna('')

    # Renombrar columnas problemáticas temporalmente
    data_temp = data_temp.rename(columns={
        'Vía pública': 'Via publica',
    })
    
    # Definir los campos obligatorios
    required_fields = [
        "Name of the Petitioning Organization", "In Care Of Name (if any) last", 
    ]
    
    for index, row in data_temp.iterrows():
        
        # Verificar campos obligatorios
        missing_fields = [field for field in required_fields if pd.isna(row[field]) or row[field] == '']
        
        if missing_fields:
            continue
        
        # Buscar información del abogado correspondiente
        attorney_row = attorney_info[attorney_info['Name of the Petitioning Organization'] == row['Name of the Petitioning Organization']]
        
        if attorney_row.empty:
            continue
        
        attorney_row = attorney_row.iloc[0]
        
        # Nombre del archivo PDF de salida
        output_pdf = f'{output_folder}/Modelo_G28i_{str(row["Name of the Petitioning Organization"])}_{str(row["Middle Name"])} {str(row["Family Name (Last Name)"])}, {str(row["Given Name (First Name)"])}.pdf'

        # Crear un lienzo PDF
        c = canvas.Canvas(output_pdf, pagesize=letter)
        
        # Crear el JSON de los datos de texto
        texto_data_hoja1 = [
            #Trabajador ----->
            {"text": str(attorney_row['Family Name (Last Name)']), "x": 121, "y": 593},
            {"text": str(attorney_row['Given Name (First Name)']), "x": 121, "y": 570},
            {"text": str(attorney_row['Middle Name']), "x": 121, "y": 545},
            {"text": str(attorney_row['Street Number and Name']), "x": 121, "y": 492},
            {"text": str(attorney_row['Apt/Ste/Flr_whiteSpace']), "x": 187, "y": 468},
            {"text": str(attorney_row['City or Town']), "x": 121, "y": 445},
            {"text": str(attorney_row['Province']), "x": 121, "y": 420},
            {"text": str(attorney_row['Postal Code']), "x": 121, "y": 395},
            {"text": str(attorney_row['Country']), "x": 61, "y": 360},
            {"text": str(attorney_row['Mobile Telephone Number (if any)']), "x": 61, "y": 257},
            {"text": str(attorney_row['Email Address (if any)']), "x": 61, "y": 220},
            {"text": str(attorney_row['Licensing Authority']), "x": 343, "y": 552},
            {"text": str(attorney_row['License Number (if applicable)']), "x": 343, "y": 515},
            
        ]
        
        if(str(attorney_row['Apt-Ste-Flr']) == "Flr"):  
            texto_data_hoja1.append({"text": "X", "x": 145, "y": 468})
        elif(str(attorney_row['Apt-Ste-Flr']) == "Ste"):
            texto_data_hoja1.append({"text": "X", "x": 103, "y": 468})
        elif(str(attorney_row['Apt-Ste-Flr']) == "Apt"):
            texto_data_hoja1.append({"text": "X", "x": 61, "y": 468})
            
            
        texto_data_hoja1.append({"text": "X", "x": 345, "y": 645})
        
        texto_data_hoja2 = [
            {"text": str(row['U.S. Street address']), "x": 404, "y": 648},
            {"text": str(row['City_Petitioner']), "x": 404, "y": 600},
            {"text": str(row['State_Petitioner']), "x": 404, "y": 576},
            {"text": str(row['Zip Code_Petitioner']), "x": 404, "y": 552},
            
            {"text": str(row['Given Name (First Name)']), "x": 121, "y": 462},
            {"text": str(row['Family Name (Last Name)']), "x": 121, "y": 437},            
            {"text": str(row['Middle Name']), "x": 121, "y": 412},
            
            {"text": str(row['Name of the Petitioning Organization']), "x": 61, "y": 376},            
            {"text": str(row['Job Title_Act']), "x": 61, "y": 342},
        ]
        
        texto_data_hoja3 = [
            # Aquí puedes agregar datos de texto adicionales si es necesario
        ]
        
        # Añadir texto a las páginas
        utils.add_text_to_image(c, image_paths[0], texto_data_hoja1)
        c.showPage()  # Añadir nueva página para la siguiente imagen
        
        utils.add_text_to_image(c, image_paths[1], texto_data_hoja2)
        c.showPage()  # Añadir nueva página para la siguiente imagen
        
        utils.add_text_to_image(c, image_paths[2], texto_data_hoja3)
        c.showPage()  # Añadir nueva página para la siguiente imagen
        
        texto_data_hoja4 = [            
            {"text": str(attorney_row['Family Name (Last Name)']), "x": 121, "y": 612},
            {"text": str(attorney_row['Given Name (First Name)']), "x": 121, "y": 589},
            {"text": str(attorney_row['Middle Name']), "x": 121, "y": 564},
        ]
        
        utils.add_text_to_image(c, image_paths[3], texto_data_hoja4)
        c.showPage()  # Añadir nueva página para la siguiente imagen
        
        # Guardar el PDF
        c.save()

        # Añadir la firma con transparencia usando PyMuPDF
        pdf_document = fitz.open(output_pdf)
        firma_path = f'{ruta_base_template}Firmas/{str(attorney_row["Signature Attorney"])}'
        page = pdf_document[2]  # La página 3 en índice 2

        # Abrir la imagen de la firma y redimensionarla
        firma_img = Image.open(firma_path)
        firma_img_resized = firma_img.resize((430, 160), Image.Resampling.LANCZOS)

        # Guardar la imagen redimensionada en memoria
        import io
        firma_img_bytes = io.BytesIO()
        firma_img_resized.save(firma_img_bytes, format='PNG')
        firma_img_bytes.seek(0)
        
        # Calcular la posición basada en las dimensiones de la imagen redimensionada
        x = 60
        y = 128
        rect = fitz.Rect(x, y, x + 160, y + 80)
        
        temp_output_pdf = f'{output_folder}/temp_Modelo_G28i_{str(row["Name of the Petitioning Organization"])}_{str(row["Middle Name"])} {str(row["Family Name (Last Name)"])}, {str(row["Given Name (First Name)"])}.pdf'
        
        page.insert_image(rect, stream=firma_img_bytes)
        pdf_document.save(temp_output_pdf)

        # Reemplazar el archivo original por el temporal
        os.replace(temp_output_pdf, output_pdf)

        # Mover la fila de Temp a Log
        row_filtered = filter_all_na_columns(pd.DataFrame([row]))  # Filter out all-NA columns
        data_log = pd.concat([data_log, row_filtered], ignore_index=True)
        data_temp = data_temp.drop(index)

# Variables de entrada
excel_file = ruta_base + 'INPUT USERS DATA FORM I-129S.xlsx'
attorney_file = ruta_base_template + 'attorney_info.xlsx'
image_paths = [
    ruta_base_template + 'page_1.jpg', 
    ruta_base_template + 'page_2.jpg',
    ruta_base_template + 'page_3.jpg',
    ruta_base_template + 'page_4.jpg'
]

# Crear la carpeta de salida si no existe
os.makedirs(output_folder, exist_ok=True)

# Ejecutar la función para generar PDFs
g28_generate_pdfs_from_excel(excel_file, attorney_file, image_paths, ruta_base_template)
