import os
from exportFinalReport import get_final_report
from flask import (Flask, redirect, render_template, request,
                   send_from_directory, url_for)
from exportFinalReport import *
from flask_executor import Executor
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from datetime import date
import threading

app = Flask(__name__)

current_date = date.today()

def run_export(file_path, proName):
    
    while not os.path.isdir(file_path):
        time.sleep(1) # wait for 1 second before checking again
    get_final_report(proName)

## powerapp only allow GET method
@app.route('/export', methods=['GET'])
def exportReport():
    param1 = request.args.get('p1', "default_projectID")
    param2 = request.args.get('p2', "default_projectName")
    param3 = request.args.get('p3', "default_userName")
    path = f'\\\\eeazurefilesne.file.core.windows.net\\generalshare\\Ethos Digital\\BMS Points Generator Reports\\Points Schedule - {param2} - {current_date.day:02d}-{current_date.month:02d}-{current_date.year}.csv'
    # pass the parameters to the ExampleHandler
    thread = threading.Thread(target=run_export, args=(path, param2))
    thread.start()
    thread.join() # main thread relies on the result of other threads, so join needed
    return render_template('index.html', param1=param1, param2=param2, param3=param3)

# @app.route('/')
# def index():
#    print('Request for index page received')
#    return render_template('index.html')

@app.route('/favicon.ico')
def favicon():
    return send_from_directory(os.path.join(app.root_path, 'static'),
                               'favicon.ico', mimetype='image/vnd.microsoft.icon')

@app.route('/hello', methods=['POST'])
def hello():
   name = request.form.get('name')
   get_final_report(name)

   if name:
       print('Request for hello page received with name=%s' % name)
       return render_template('hello.html', name = name)
   else:
       print('Request for hello page received with no name or blank name -- redirecting')
       return redirect(url_for('index'))



if __name__ == '__main__':
   app.run()