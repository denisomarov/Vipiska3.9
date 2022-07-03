# import sqlite3
import os
from flask import Flask, render_template, url_for, request
from treatment import v_treatment

# конфигурация
DATABASE = '/tmp/flsite.db'
DEBUG = True
SECRET_KEY = 'sfsfnkl5$klm^ds#m0>sdf3_9'

app = Flask(__name__)
app.config.from_object(__name__)
app.config['SECRET_KEY'] = 'cvknxcvlksnvlncvvxcnvlxjcnv'

menu = [{"name": "Установка", "url": "install-flask"},
        {"name": "Первое приложение", "url": "first-app"},
        {"name": "Обратная связь", "url": "contact"}]

@app.route('/', methods=["POST", "GET"])
def index():

    if request.method == 'POST':
        Halykfile = request.files['Halykfileupload']
        HalykFileName = os.path.abspath(Halykfile.filename)

        Sberfile = request.files['Sberfileupload']
        SberFileName = os.path.abspath(Sberfile.filename)

        if os.path.isfile(HalykFileName) and os.path.isfile(SberFileName):
            v_treatment(HalykFileName, SberFileName)

    return render_template('index.html', title='Обработка выписок из банка', menu=menu)

if __name__ == '__main__':
    app.run()
