#pip install -r requirements.txt
from flask import Flask, render_template, request, redirect, url_for, flash
from Bot_247 import bot247
from Bot_i129s import read_excels, bot_i129s
import logging

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # Add a secret key for session management
logging.basicConfig(level=logging.INFO)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/start_bot247', methods=['POST'])
def start_bot247():
    try:
        bot247_excel_file = 'G:/Shared drives/ES VIALTO GMS - RPA/TAX/COMPLIANCE/247/datos.xlsx'
        bot247_image_paths = [
            'G:/Shared drives/ES VIALTO GMS - RPA/TAX/COMPLIANCE/247/templates/page_1.png', 
            'G:/Shared drives/ES VIALTO GMS - RPA/TAX/COMPLIANCE/247/templates/page_2.png'
        ]
        bot247_output_folder = 'G:/Shared drives/ES VIALTO GMS - RPA/TAX/COMPLIANCE/247/pdfs_generados'
        bot247.generate_pdfs_from_excel(bot247_excel_file, bot247_image_paths, bot247_output_folder)
        flash('Bot 247 completed successfully', 'success')
    except Exception as e:
        logging.error(f"Error in start_bot247: {e}")
        flash('An error occurred while running Bot 247', 'danger')
    return redirect(url_for('index'))

@app.route('/start_i129s', methods=['POST'])
def start_i129s():
    try:
        bot_i129s_excel_file = 'G:/Shared drives/ES VIALTO GMS - RPA/TAX/COMPLIANCE/i_129s/datos.xlsx'
        bot_i129s_image_paths = [
            'G:/Shared drives/ES VIALTO GMS - RPA/TAX/COMPLIANCE/i_129s/templates/page_1.jpg', 
            'G:/Shared drives/ES VIALTO GMS - RPA/TAX/COMPLIANCE/i_129s/templates/page_2.jpg',
            'G:/Shared drives/ES VIALTO GMS - RPA/TAX/COMPLIANCE/i_129s/templates/page_3.jpg',
            'G:/Shared drives/ES VIALTO GMS - RPA/TAX/COMPLIANCE/i_129s/templates/page_4.jpg',
            'G:/Shared drives/ES VIALTO GMS - RPA/TAX/COMPLIANCE/i_129s/templates/page_5.jpg',
            'G:/Shared drives/ES VIALTO GMS - RPA/TAX/COMPLIANCE/i_129s/templates/page_6.jpg',
            'G:/Shared drives/ES VIALTO GMS - RPA/TAX/COMPLIANCE/i_129s/templates/page_7.jpg',
            'G:/Shared drives/ES VIALTO GMS - RPA/TAX/COMPLIANCE/i_129s/templates/page_8.jpg'
        ]
        bot_i129s_output_folder = 'G:/Shared drives/ES VIALTO GMS - RPA/TAX/COMPLIANCE/i_129s/pdfs_generados'
        bot_i129s.generate_pdfs_from_excel(bot_i129s_excel_file, bot_i129s_image_paths, bot_i129s_output_folder)
        flash('Bot i129s completed successfully', 'success')
    except Exception as e:
        logging.error(f"Error in start_i129s: {e}")
        flash('An error occurred while running Bot i129s. Check the excel', 'danger')
    return redirect(url_for('index'))

@app.route('/start_excel_i129s', methods=['POST'])
def start_excel_i129s():
    try:
        read_excels.generateExcel()
        flash('Excel generation completed successfully', 'success')
    except Exception as e:
        logging.error(f"Error in start_excel_i129s: {e}")
        flash('An error occurred while generating the Excel file', 'danger')
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)