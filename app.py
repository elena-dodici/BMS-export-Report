import os
from exportFinalReport import get_final_report
from flask import (Flask, redirect, render_template, request,
                   send_from_directory, url_for)
from exportFinalReport import *


app = Flask(__name__)


## powerapp only allow GET method
@app.route('/export', methods=['GET'])
def exportReport():
    param1 = request.args.get('p1', "default_value1")
    param2 = request.args.get('p2', "default_value2")
    param3 = request.args.get('p3', "default_value3")
  
    get_final_report(param2)
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