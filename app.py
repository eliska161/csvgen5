from flask import Flask, request, send_file
import pandas as pd
from openpyxl import load_workbook

app = Flask(__name__)

@app.route('/')
def upload_page():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process_files():
    # Mottar CSV-filen
    csv_file = request.files['csv_file']

    # Laste inn den statiske Excel-malen som er lagret i serveren (eks. assets/mal.xlsx)
    excel_path = 'assets/Regnskapsark-for-elevbedrifter.xlsx'  # Bruk din egen mal her
    wb = load_workbook(excel_path)
    ws = wb['Ark1']

    # Les CSV-data
    csv_data = pd.read_csv(csv_file)
    csv_data['Dato'] = pd.to_datetime(csv_data['Dato']).dt.strftime('%d.%m.%Y')
    csv_data['Ut'] = csv_data['Beløp'].apply(lambda x: -x if x < 0 else 0)
    csv_data['Inn'] = csv_data['Beløp'].apply(lambda x: x if x > 0 else 0)

    # Sett data inn i Excel-malen
    for i, row in csv_data.iterrows():
        if 6 + i > 19:  # Bare til A19
            break
        ws[f"A{6 + i}"] = f"A{6 + i}"  # Billag
        ws[f"B{6 + i}"] = row['Dato']
        ws[f"C{6 + i}"] = row['Beskrivelse']
        ws[f"D{6 + i}"] = row['Ut']
        ws[f"E{6 + i}"] = row['Inn']
    
    output_path = "Updated_File.xlsx"
    wb.save(output_path)

    return send_file(output_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
