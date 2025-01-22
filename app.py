from flask import Flask, request, send_file, render_template
import pandas as pd
from openpyxl import load_workbook

app = Flask(__name__)

@app.route('/')
def upload_page():
    return render_template('index.html')  # En enkel opplastingsside

@app.route('/process', methods=['POST'])
def process_files():
    # Hent filene fra opplastingen
    csv_file = request.files['csv_file']
    excel_file = request.files['excel_file']
    
    # Les CSV og Excel
    csv_data = pd.read_csv(csv_file)
    csv_data['Dato'] = pd.to_datetime(csv_data['Dato']).dt.strftime('%d.%m.%Y')
    csv_data['Ut'] = csv_data['Beløp'].apply(lambda x: -x if x < 0 else 0)
    csv_data['Inn'] = csv_data['Beløp'].apply(lambda x: x if x > 0 else 0)
    
    # Last inn Excel-filen
    wb = load_workbook(excel_file)
    ws = wb['Ark1']  # Velg "Ark1"

    # Sett data i spesifikke celler
    for i, row in csv_data.iterrows():
        if 6 + i > 19:  # Forhindrer overskriving utenfor A6-E19
            break
        ws[f"A{6 + i}"] = f"A{6 + i}"  # Billag
        ws[f"B{6 + i}"] = row['Dato']
        ws[f"C{6 + i}"] = row['Beskrivelse']
        ws[f"D{6 + i}"] = row['Ut']
        ws[f"E{6 + i}"] = row['Inn']
    
    # Lagre oppdatert Excel-fil
    output_path = "Updated_File.xlsx"
    wb.save(output_path)
    
    return send_file(output_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, port=8080)
