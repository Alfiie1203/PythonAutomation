import os
from reportlab.lib.pagesizes import letter
from reportlab.lib import utils
from reportlab.pdfgen import canvas
import pandas as pd

output_folder = 'G:/Shared drives/ES VIALTO GMS - RPA/TAX/COMPLIANCE/247/pdfs_generados'

# Create folder if not exist
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

''' 
Load the images
Scale the images like A4 page
Add images
Adjust text size
'''
def add_text_to_image(canvas_obj, image_path, text_data):
    # Cargar la imagen
    img = utils.ImageReader(image_path)
    img_width, img_height = img.getSize()

    # Escalar la imagen para que quepa en la página
    scaling_factor = min(letter[0] / img_width, letter[1] / img_height)
    img_width *= scaling_factor
    img_height *= scaling_factor

    # Calcular la posición para centrar la imagen en la página
    x_pos = (letter[0] - img_width) / 2
    y_pos = (letter[1] - img_height) / 2

    # Agregar la imagen al lienzo
    canvas_obj.drawImage(image_path, x_pos, y_pos, width=img_width, height=img_height)

    # Ajustar el tamaño del texto
    canvas_obj.setFont("Helvetica", 7)  # Cambia 7 al tamaño de fuente deseado

    # Agregar texto desde el JSON
    for text_entry in text_data:
        text = text_entry['text']
        text_x = text_entry['x']
        text_y = text_entry['y']
        # Agregar texto al lienzo en la posición especificada
        canvas_obj.drawString(text_x, text_y, text)


# Converts the zip code to a string with spaces between the digits
def addSpaces(varSpaces):
    return ' '.join(varSpaces)

# Convert the date format from 'YYYY-DD-MM HH:MM:SS' to 'DDMMYYYY'
def convert_date_format(date_str): 
    date_obj = pd.to_datetime(date_str)
    return date_obj.strftime('%d%m%Y')

'''
Read the excel file
Temporarily rename problematic columns
Create the JSON of the text data
Add text and image to each page
Move the row from Temp to Log
Save changes to the Excel file
'''
def generate_pdfs_from_excel(excel_file, image_paths, output_folder):
    # Leer el archivo Excel
    xls = pd.ExcelFile(excel_file)
    data_temp = pd.read_excel(xls, 'Temp')
    data_log = pd.read_excel(xls, 'Log')

    # Renombrar columnas problemáticas temporalmente
    data_temp = data_temp.rename(columns={
        'Vía pública': 'Via publica',
        'Núm.': 'Num.',
        'COMPAÑIA': 'Compania',  
        'Código Postal': 'Codigo_Postal',
        'Fecha de salida del territorio español': 'Fecha de salida',
        'Fecha de comienzo de la prestación del trabajo en el otro país': 'Fecha de comienzo',
        'País o territorio de desplazamiento' : 'Pa_desplazamiento'
    })

    for index, row in data_temp.iterrows():
        # Crear el JSON de los datos de texto
        texto_data = [
            #Trabajador ----->
            {"id": 1, "text": str(row['NIF']), "x": 91, "y": 577},
            {"id": 2, "text": str(row['Primer apellido']), "x": 200, "y": 577},
            {"id": 3, "text": str(row['Segundo apellido']), "x": 111, "y": 567},
            {"id": 4, "text": str(row['Nombre']), "x": 250, "y": 567},
            {"id": 5, "text": str(row['Via publica']), "x": 100, "y": 545},
            {"id": 6, "text": str(row['Num.']), "x": 260, "y": 545},
            {"id": 7, "text": str(row['Esc.']), "x": 285, "y": 545},
            {"id": 8, "text": str(row['Piso']), "x": 307, "y": 545},
            {"id": 9, "text": str(row['Prta.']), "x": 328, "y": 545},
            {"id": 10, "text": str(row['Municipio']), "x": 97, "y": 535},
            {"id": 11, "text": str(row['Provincia']), "x": 210, "y": 535},
            {"id": 12, "text": addSpaces(str(row['Codigo_Postal'])), "x": 312, "y": 535},
            #Domicilio para notificaciones ----->
            #{"id": 11, "text": str(row['Provincia']), "x": 210, "y": 535},
            #Datos identificativos del pagador de los rendimientos del trabajo ----->
            {"id": 13, "text": str(row['CIF']), "x": 95, "y": 409},
            {"id": 14, "text": str(row['Compania']), "x": 208, "y": 409},
            {"id": 15, "text": str(row['Domicilio fiscal']), "x": 95, "y": 385},
            #Datos identificativos del pagador de los rendimientos del trabajo ----->
            {"id": 16, "text": addSpaces(addSpaces(convert_date_format(row['Fecha de salida']))), "x": 355, "y": 308},
            {"id": 17, "text": addSpaces(addSpaces(convert_date_format(row['Fecha de comienzo']))), "x": 355, "y": 290},
            {"id": 18, "text": row['Pa_desplazamiento'], "x": 355, "y": 272},
        ]
       
        if(str(row['Domicilio para notificaciones']) == "VIALTO"):  
            #Domicilio para notificaciones ----->
            texto_data.append({"id": 19, "text": "Vialto Partners Spain SLU", "x": 362, "y": 549})
            texto_data.append({"id": 20, "text": "Paseo de Recoletos", "x": 381, "y": 527})
            texto_data.append({"id": 20, "text": "5", "x": 444, "y": 517})
            texto_data.append({"id": 21, "text": "", "x": 468, "y": 517})
            texto_data.append({"id": 22, "text": "4", "x": 493, "y": 517})
            texto_data.append({"id": 23, "text": "", "x": 513, "y": 517})
            texto_data.append({"id": 24, "text": "Madrid", "x": 380, "y": 507})
            texto_data.append({"id": 25, "text": "Madrid", "x": 380, "y": 496})
            texto_data.append({"id": 26, "text": addSpaces("28004"), "x": 503, "y": 497})
           
        elif(str(row['Domicilio para notificaciones']) == "PROPIO"):
            #Domicilio para notificaciones ----->
            texto_data.append({"id": 19, "text": str(row['Nombre']+ " "+ str(row['Primer apellido']) + " "+str(row['Segundo apellido'])), "x": 362, "y": 549})
            texto_data.append({"id": 20, "text": str(row['Via publica']), "x": 380, "y": 527})
            texto_data.append({"id": 20, "text": str(row['Num.']), "x": 444, "y": 517})
            texto_data.append({"id": 21, "text": str(row['Esc.']), "x": 468, "y": 517})
            texto_data.append({"id": 22, "text": str(row['Piso']), "x": 493, "y": 517})
            texto_data.append({"id": 23, "text": str(row['Prta.']), "x": 513, "y": 517})
            texto_data.append({"id": 24, "text": str(row['Municipio']), "x": 380, "y": 507})
            texto_data.append({"id": 25, "text": str(row['Provincia']), "x": 380, "y": 496})
            texto_data.append({"id": 26, "text": addSpaces(str(row['Codigo_Postal'])), "x": 503, "y": 497})          

        # Nombre del archivo PDF de salida
        output_pdf = f'{output_folder}/Modelo_247_{str(row['Segundo apellido'])} {str(row['Primer apellido'])}, {str(row['Nombre'])}.pdf'

        # Crear un lienzo PDF
        c = canvas.Canvas(output_pdf, pagesize=letter)

        # Agregar texto e imagen a cada página
        for image_path in image_paths:
            add_text_to_image(c, image_path, texto_data)
            c.showPage()  # Añadir nueva página para la siguiente imagen

        # Guardar el PDF
        c.save()

        # Mover la fila de Temp a Log
        data_log = pd.concat([data_log, pd.DataFrame([row])])
        data_temp = data_temp.drop(index)

    # Guardar los cambios en el archivo Excel
    with pd.ExcelWriter(excel_file, mode='a', if_sheet_exists='replace') as writer:
        data_temp.to_excel(writer, sheet_name='Temp', index=False)
        data_log.to_excel(writer, sheet_name='Log', index=False)


#excel_file = 'G:/Shared drives/ES VIALTO GMS - RPA/TAX/COMPLIANCE/247/datos.xlsx'
#image_paths = ['G:/Shared drives/ES VIALTO GMS - RPA/TAX/COMPLIANCE/247/templates/page_1.png', 
#               'G:/Shared drives/ES VIALTO GMS - RPA/TAX/COMPLIANCE/247/templates/page_2.png']
#output_folder = 'G:/Shared drives/ES VIALTO GMS - RPA/TAX/COMPLIANCE/247/pdfs_generados'
#generate_pdfs_from_excel(excel_file, image_paths, output_folder)
