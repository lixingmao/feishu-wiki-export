from flask import Flask
from playground.data import process_data
from playground.template import upload_file, download_file

app = Flask(__name__)

@app.route('/')
def index():
    return 'Feishu wiki base export script connected successfully!'

@app.route('/download_template')
def download_template():
    return 'Coming soon!'

@app.route('/upload_template')
def upload_template():
    return upload_file()

@app.route('/process_data')
def process_data():
    return process_data()


app.run(host='0.0.0.0', port=81)
