import requests
from flask import Flask, request, send_file, render_template
import os

app = Flask(__name__)

# Worker URL hvor du vil sende CSV-filen og malen for prosessering
WORKER_URL = "https://csvworker.elias-8cc.workers.dev"  # Sett URL-en til din Cloudflare Worker

@app.route('/')
def upload_page():
    return render_template('index.html')  # Sender brukeren til opplastingsskjemaet

@app.route('/process', methods=['POST'])
def upload_files():
    # Sjekk om både CSV-filen og Excel-malen er lastet opp
    csv_file = request.files['csv_file']
    excel_template = os.path.join('assets', 'Regnskapsark-for-elevbedrifter.xlsx')  # Stien til Excel-malen
    
    if not csv_file:
        return "Feil: CSV-fil må lastes opp", 400
    
    if not os.path.exists(excel_template):
        return "Feil: Excel-mal ikke funnet", 500

    # Åpne og send filene til Cloudflare Worker
    with open(excel_template, 'rb') as template_file:
        files = {
            'csv_file': (csv_file.filename, csv_file.stream, csv_file.content_type),
            'excel_template': ('Regnskapsmal.xlsx', template_file, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        }

        try:
            # Send CSV-filen og Excel-malen til Workeren for behandling
            response = requests.post(WORKER_URL, files=files)

            if response.status_code == 200:
                # Hvis Workeren returnerte filen, send den tilbake til brukeren
                return send_file(response.content, as_attachment=True, download_name="updated-regnskapsark.xlsx")
            else:
                return f"Feil med behandling på Workeren: {response.status_code}", 500
        except requests.exceptions.RequestException as e:
            return f"Kunne ikke sende forespørsel til Worker: {str(e)}", 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8080)
