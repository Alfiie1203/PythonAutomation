#pip install -r requirements.txt
from flask import Flask, render_template, request, redirect, url_for, flash
from Bot_247 import bot247
from Bot_i129s import read_excels, bot_i129s, generate_ExcelKey
import logging

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # Add a secret key for session management
logging.basicConfig(level=logging.INFO)

RPA = 'G:/Shared drives/ES VIALTO GMS - RPA/'
TAX_247 = RPA+'TAX/COMPLIANCE/247/'
TAX_Templates= TAX_247+'templates/'

INMI = 'TAX/COMPLIANCE/i_129s/'
INMI_I129S= INMI+'TAX/COMPLIANCE/i_129s/'
INMI_Templates=INMI_I129S+'templates/'

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/start_bot247', methods=['POST'])
def start_bot247():
    try:
        bot247_excel_file = TAX_247+'datos.xlsx'
        bot247_image_paths = [
            TAX_Templates+'page_1.png',
            TAX_Templates+'page_2.png'
        ]
        bot247_output_folder = TAX_247+'pdfs_generados'
        bot247.generate_pdfs_from_excel(bot247_excel_file, bot247_image_paths, bot247_output_folder)
        flash('Bot 247 completed successfully', 'success')
    except Exception as e:
        logging.error(f"Error in start_bot247: {e}")
        flash('An error occurred while running Bot 247', 'danger')
    return redirect(url_for('index'))

@app.route('/start_i129s', methods=['POST'])
def start_i129s():
    try:
        bot_i129s_excel_file = INMI_I129S+'datos.xlsx'
        bot_i129s_image_paths = [
            INMI_Templates+'page_1.jpg', 
            INMI_Templates+'page_2.jpg',
            INMI_Templates+'page_3.jpg',
            INMI_Templates+'page_4.jpg',
            INMI_Templates+'page_5.jpg',
            INMI_Templates+'page_6.jpg',
            INMI_Templates+'page_7.jpg',
            INMI_Templates+'page_8.jpg'
        ]
        bot_i129s_output_folder = INMI_I129S+'pdfs_generados'
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


@app.route('/start_bot_generateExcel_i129s', methods=['POST'])
def start_bot():
    try:
        read_excels.start_bot()
        bot_i129s.start_bot()
        flash('Bot started successfully', 'success')
    except Exception as e:
        logging.error(f"Error in start_bot: {e}")
        flash('An error occurred while starting the bot', 'danger')
    return redirect(url_for('index'))


@app.route('/stop_bot_generateExcel_i129s', methods=['POST'])
def stop_bot():
    try:
        read_excels.stop_bot()
        bot_i129s.stop_bot()
        flash('Bot stopped successfully', 'success')
    except Exception as e:
        logging.error(f"Error in stop_bot: {e}")
        flash('An error occurred while stopping the bot', 'danger')
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)