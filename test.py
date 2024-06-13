from flask import Flask, request, jsonify, send_file, render_template_string
import logging
import openpyxl
from openpyxl import Workbook
import os
app = Flask(__name__)
logging.basicConfig(level=logging.INFO)
excel_file = "data.xlsx"
def create_excel_file():
    if not os.path.exists(excel_file):
        wb = Workbook()
        ws = wb.active
        ws.append(["Millis", "P1"])
        wb.save(excel_file)
        logging.info(f"Created new Excel file: {excel_file}")

create_excel_file()

@app.route('/')
def home():
    # Générer la page HTML pour permettre le téléchargement du fichier Excel
    html = '''
    <!DOCTYPE html>
    <html>
    <head>
        <title>Data Download</title>
    </head>
    <body>
        <h1>Download Data</h1>
        <a href="/download_data" download>Download Data as Excel</a>
    </body>
    </html>
    '''
    return render_template_string(html)

@app.route('/data', methods=['POST'])
def receive_data():
    data = request.json
    if data:
        millis = data.get('millis')
        p1 = data.get('p1')
        logging.info(f'Received data - Millis: {millis}, P1: {p1}')
        # Stocker les données reçues dans le fichier Excel
        wb = openpyxl.load_workbook(excel_file)
        ws = wb.active
        ws.append([millis, p1])
        wb.save(excel_file)
        logging.info(f"Data saved to Excel file: {excel_file}")
        response = {'message': 'Data received', 'millis': millis, 'p1': p1}
        return jsonify(response)
    else:
        logging.error('No data received')
        return jsonify({'message': 'No data received'}), 400

@app.route('/download_data', methods=['GET'])
def download_data():
    try:
        logging.info(f"Attempting to send file: {excel_file}")
        return send_file(excel_file, as_attachment=True)
    except Exception as e:
        logging.error(f"Error sending file: {e}")
        return str(e), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)