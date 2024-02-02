from flask import Flask, render_template, request, redirect, url_for, send_file
import pandas as pd
from datetime import datetime

app = Flask(__name__)

vin_list = {'VIN': [], 'Timestamp': []}

@app.route('/')
def index():
    return render_template('index.html', vin_list=vin_list)

@app.route('/add_vin', methods=['POST'])
def add_vin():
    vin = request.form['vin']
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    vin_list['VIN'].append(vin)
    vin_list['Timestamp'].append(timestamp)
    return redirect(url_for('index'))

@app.route('/clear_vin')
def clear_vin():
    vin_list['VIN'].clear()
    vin_list['Timestamp'].clear()
    return redirect(url_for('index'))

@app.route('/undo_vin')
def undo_vin():
    if vin_list['VIN']:
        vin_list['VIN'].pop()
        vin_list['Timestamp'].pop()
    return redirect(url_for('index'))

@app.route('/save_excel', methods=['POST'])
def save_excel():
    df = pd.DataFrame(vin_list)
    filename= 'vin_list.xlsx'
    df.to_excel(filename, index=False)
    return send_file(filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)

