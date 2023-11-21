import os
from exportFinalReport import get_final_report
from flask import (Flask, redirect, render_template, request,
                   send_from_directory, url_for)

app = Flask(__name__)


@app.route('/')
def index():
   print('Request for index page received')
   return render_template('index.html')

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


#### IN USE
@app.route('/report', methods=['GET'])
    
def generate_report_with_par():
    name = request.form.get('name')
    get_final_report(name)
    return render_template('hello.html', name = name)

if __name__ == '__main__':
   app.run()
