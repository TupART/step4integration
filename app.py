from flask import Flask, request, render_template, send_file
import pandas as pd
import openpyxl
from werkzeug.utils import secure_filename
import os

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Subir el archivo
        if 'file' not in request.files:
            return 'No file part'
        file = request.files['file']
        if file.filename == '':
            return 'No selected file'
        if file:
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)

            # Leer archivo subido
            df = pd.read_excel(file_path)

            # Crear una lista de checkboxes con Name y Surname
            names = df[['Name', 'Surname']].dropna()
            return render_template('index.html', names=names.values)

    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    selected_rows = request.form.getlist('rows')

    # Cargar archivo subido previamente
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], os.listdir(app.config['UPLOAD_FOLDER'])[0])
    df = pd.read_excel(file_path)

    # Cargar la plantilla 'PlantillaSTEP4.xlsx'
    plantilla = 'PlantillaSTEP4.xlsx'
    wb = openpyxl.load_workbook(plantilla)
    ws = wb.active

    # Procesar las filas seleccionadas
    for row in selected_rows:
        idx = int(row)  # Convertir el índice de string a entero

        name = df.iloc[idx, df.columns.get_loc('Name')]
        surname = df.iloc[idx, df.columns.get_loc('Surname')]
        email = df.iloc[idx, df.columns.get_loc('E-mail')]
        market = df.iloc[idx, df.columns.get_loc('Market')]
        pcc_status = df.iloc[idx, df.columns.get_loc('Va a ser PCC?')]
        b2e_username = df.iloc[idx, df.columns.get_loc('B2E User Name')]

        # Escribir los datos en las columnas correspondientes de la plantilla
        row_num = 7 + idx  # Empezando desde la fila 7 en la plantilla
        ws[f'C{row_num}'] = name
        ws[f'D{row_num}'] = surname
        ws[f'E{row_num}'] = email

        # Condiciones para "Primary phone"
        if pcc_status == 'Y' and market == 'DACH':
            ws[f'F{row_num}'] = "/+4940210918145 /+43122709858 /+41445295828"
        elif pcc_status == 'Y' and market == 'France':
            ws[f'F{row_num}'] = "/+33180037979"
        elif pcc_status == 'Y' and market == 'Spain':
            ws[f'F{row_num}'] = "/+34932952130"
        elif pcc_status == 'Y' and market == 'Italy':
            ws[f'F{row_num}'] = "/+390109997099"

        # Condiciones para "Workgroup"
        if pcc_status == 'Y' and market == 'DACH':
            ws[f'G{row_num}'] = "D_PCC"
        elif pcc_status == 'Y' and market == 'France':
            ws[f'G{row_num}'] = "F_PCC"
        elif pcc_status == 'Y' and market == 'Spain':
            ws[f'G{row_num}'] = "E_PCC"
        elif pcc_status == 'Y' and market == 'Italy':
            ws[f'G{row_num}'] = "I_PCC"

        # Condiciones para "Team"
        if pcc_status == 'Y' and market == 'DACH':
            ws[f'H{row_num}'] = "Team_D_CCH_PCC_1"
        elif pcc_status == 'Y' and market == 'France':
            ws[f'H{row_num}'] = "Team_F_CCH_PCC_1"
        elif pcc_status == 'Y' and market == 'Spain':
            ws[f'H{row_num}'] = "Team_E_CCH_PCC_1"
        elif pcc_status == 'Y' and market == 'Italy':
            ws[f'H{row_num}'] = "Team_I_CCH_PCC_1"

        # Condiciones para "Is PCC"
        if pcc_status == 'Y':
            ws[f'L{row_num}'] = "Y"
        else:
            ws[f'L{row_num}'] = "N"

        # Escribir otros campos
        ws[f'Q{row_num}'] = b2e_username
        ws[f'R{row_num}'] = b2e_username

        # Condiciones para "Campaign Level"
        if pcc_status in ['Y', 'N', 'DS']:
            ws[f'V{row_num}'] = "Agent"
        elif pcc_status == 'TL':
            ws[f'V{row_num}'] = "Team Leader"

    # Guardar la plantilla con los datos rellenados
    output_file = 'PlantillaSTEP4_Rellenada.xlsx'
    wb.save(output_file)

    # Enviar el archivo descargable
    return send_file(output_file, as_attachment=True)

if __name__ == '__main__':
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    app.run(debug=True)
