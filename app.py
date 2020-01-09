import os
from flask import Flask, render_template,url_for, request, send_file
from werkzeug.utils import secure_filename
import excelConverter

app = Flask(__name__)
UPLOAD_FOLDER='uploads'
app.config['UPLOAD_FOLDER']=UPLOAD_FOLDER

@app.route('/', host="skureshy.pythonanywhere.com")
def upload():
    return render_template('upload.html')

@app.route('/uploader', methods=['GET','POST'], host="skureshy.pythonanywhere.com")
def upload_file():
    # testing
    if request.method == 'POST':
        f = request.files['file']
        filePath = os.path.join(app.config['UPLOAD_FOLDER'],secure_filename(f.filename))
        f.save(filePath)
        spreadsheetPath = 'report'
        excelConverter.getSpreadsheet(filePath, spreadsheetPath)
        return send_file(spreadsheetPath+'.xlsx',as_attachment=True, attachment_filename=spreadsheetPath+'.xlsx')
    else:
        return upload()

@app.errorhandler(404)
def page_not_found(e):
    return render_template('404.html')

@app.errorhandler(500)
def page_not_found(e):
    return render_template('500.html')

if __name__ == '__main__':
    app.run(debug=True)
