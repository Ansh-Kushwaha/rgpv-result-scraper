import os
import requests
import pytesseract
import random
import string
from io import BytesIO, StringIO
from PIL import Image
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor
from time import sleep
import threading
import csv

from flask import Flask, render_template, request, jsonify, send_file, Response, url_for
from flask_socketio import SocketIO

app = Flask(__name__)
socketio = SocketIO(app)

# Configure Tesseract path
pytesseract.pytesseract.tesseract_cmd = 'C:/Program Files/Tesseract-OCR/tesseract'

def get_random_string():
    random_str = ''.join([random.choice(string.ascii_letters + string.digits) for _ in range(24)])
    return random_str

class Processor:
    fail = False
    processed_count = 0
    total_count = 0

    def __init__(self, sem, branch=""):
        self.lock = threading.Lock()
        self.sem = sem
        self.branch = branch
        self.first_entry = True
        self.results = {}
        self.num_cols = False

    def start(self, first_roll, last_roll):
        self.roll_list = self.roll_list_generator(first_roll, last_roll)
        self.total_count = len(self.roll_list)
        sess_url = self.get_session()
        if self.fail:
            return False
        self.sess, self.url = sess_url
        self.process(wait=True)
        return True

    def process(self, wait=False):
        with ThreadPoolExecutor(max_workers=200) as executor:
            for roll in self.roll_list:
                executor.submit(self.try_open, roll)
            executor.shutdown(wait=wait)

    def try_open(self, roll):
        while self.get_result(roll) == 1:
            pass

    def get_session(self):
        try:
            cookie = get_random_string()
            header = {
                'User-Agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Mobile Safari/537.36 Edg/130.0.0.0',
                'Cookies': 'ASP.NET_SessionId=' + cookie
            }

            sess = requests.session()
            sess.headers.update(header)
            program_resp = sess.get('http://result.rgpv.ac.in/Result/ProgramSelect.aspx')

            soup = BeautifulSoup(program_resp.text, 'html5lib')

            deptid = 'radlstProgram_1'
            value = soup.find('input', {'id': deptid})['value']
            deptid = deptid.replace('_', '$')
            viewState = soup.find('input', {'id': '__VIEWSTATE'})['value']
            viewStateGen = soup.find('input', {'id': '__VIEWSTATEGENERATOR'})['value']
            EvenValidation = soup.find('input', {'id': '__EVENTVALIDATION'})['value']
            post_data = {
                '__EVENTTARGET': deptid,
                '__EVENTARGUMENT': '',
                '__LASTFOCUS': '',
                '__VIEWSTATE': viewState,
                '__VIEWSTATEGENERATOR': viewStateGen,
                '__EVENTVALIDATION': EvenValidation,
                'radlstProgram': value
            }
            resp = sess.post('http://result.rgpv.ac.in/Result/ProgramSelect.aspx', data=post_data, allow_redirects=True)
            url = resp.url
            return (sess, url)

        except Exception as e:
            print("Exception while establishing session: ", e)
            socketio.emit('error', {'message': f"Session error: {str(e)}"})
            self.fail = True

    def get_result(self, roll):
        for _ in range(10):
            try:
                with self.lock:
                    resp = self.sess.get(self.url)
                
                soup = BeautifulSoup(resp.text, 'html5lib')
                image_url = "http://result.rgpv.ac.in/Result/" + soup.findAll('img')[1]['src']
                response = requests.get(image_url)
                
                if response.status_code != 200:
                    return 1
                
                img = Image.open(BytesIO(response.content))
                solution = pytesseract.image_to_string(img).strip().upper().replace(' ', '')
                if not solution:
                    return 1
                
                # minimum 5 second delay
                sleep(5)
                viewState = soup.find('input', {'id': '__VIEWSTATE'})['value']
                viewStateGen = soup.find('input', {'id': '__VIEWSTATEGENERATOR'})['value']
                EvenValidation = soup.find('input', {'id': '__EVENTVALIDATION'})['value']

                post_data = {
                    '__EVENTTARGET': '',
                    '__EVENTARGUMENT': '',
                    '__LASTFOCUS': '',
                    '__VIEWSTATE': viewState,
                    '__VIEWSTATEGENERATOR': viewStateGen,
                    '__EVENTVALIDATION': EvenValidation,
                    'ctl00$ContentPlaceHolder1$txtrollno': roll,
                    'ctl00$ContentPlaceHolder1$drpSemester': str(self.sem),
                    'ctl00$ContentPlaceHolder1$rbtnlstSType': 'G',
                    'ctl00$ContentPlaceHolder1$TextBox1': solution,
                    'ctl00$ContentPlaceHolder1$btnviewresult': 'View Result'
                }

                with self.lock:
                    result = self.sess.post(self.url, data=post_data, allow_redirects=True)

                result_found = '<td class="resultheader">'
                wrong_captcha = '<script language="JavaScript">alert("you have entered a wrong_captcha text");</script>'
                result_not_found = '<script language=JavaScript>alert("Result for this Enrollment No. not Found");</script>'
                
                if result_found in result.text:
                    self.process_result(result.text, roll)
                    return 0
                
                elif wrong_captcha in result.text:
                    return 1
                elif result_not_found in result.text:
                    with self.lock:
                        self.processed_count += 1
                        progress = int((self.processed_count / self.total_count) * 100)
                        socketio.emit('progress', {'count': self.processed_count, 'total': self.total_count, 'percent': progress})
                    return 0
                else:
                    return 1

            except Exception as e:
                print("Exception while opening result for roll: ", roll, e)
                socketio.emit('error', {'message': f"Error processing {roll}: {str(e)}"})
                self.fail = True
            else:
                break
        else:
            self.fail = True

    def process_result(self, html, roll):
        list_data = []
        soup = BeautifulSoup(html, 'html5lib')

        name = soup.find(id="ctl00_ContentPlaceHolder1_lblNameGrading").get_text().strip()
        sgpa = soup.find(id="ctl00_ContentPlaceHolder1_lblSGPA").get_text()
        cgpa = soup.find(id="ctl00_ContentPlaceHolder1_lblcgpa").get_text()
        result = soup.find(id="ctl00_ContentPlaceHolder1_lblResultNewGrading").get_text()
        
        with self.lock:
            list_data.append(roll)
            list_data.append(name)
        
        results = soup.findAll("table")[0].findAll("table")[2].findAll("tr")[6].findAll("table")
        
        with self.lock:
            if self.first_entry is True:
                self.first_entry = False
                header_row = []
                header_row.append("Enrollment Number")
                header_row.append("Name")
                for row in range(1, len(results)):
                    header_row.append(results[row].findAll('td')[0].text.replace("\n", '').strip())
                header_row.append("SGPA")
                header_row.append("CGPA")
                header_row.append("Result")
                self.results[0] = header_row
                self.num_cols = len(header_row)

        for row in range(1, len(results)):
            list_data.append(results[row].findAll('td')[3].text.replace("\n", '').strip())

        list_data.append(sgpa)
        list_data.append(cgpa)
        list_data.append(result)
        
        with self.lock:
            self.processed_count += 1
            self.results[int(roll[-3:])] = list_data
            
            progress = int((self.processed_count / self.total_count) * 100)
            socketio.emit('progress', {'count': self.processed_count, 'total': self.total_count, 'percent': progress})

    def to_csv(self):
        if self.fail:
            return None
        
        sorted_list = sorted(self.results.items())
        if self.num_cols:
            # Convert to CSV string
            output = StringIO()
            writer = csv.writer(output)
            for _, row in sorted_list:
                writer.writerow(row)
            return output.getvalue()
        else:
            return None

    def roll_list_generator(self, first_roll, last_roll):
        if len(first_roll) != len(last_roll):
            socketio.emit('error', {'message': "Incorrect enrollment numbers format."})
            return []

        roll_list = []
        start = int(first_roll[-4:])
        end = int(last_roll[-4:]) + 1
        common = first_roll[:8]
        for i in range(start, end):
            i = str(i).zfill(4)  # Zero-pad to ensure consistent length
            roll = common + i
            roll_list.append(roll)
        
        return roll_list

# Flask routes
@app.route('/')
def index():
    branches = [
        {"code": "CS", "name": "Computer Science Engineering"},
        {"code": "AD", "name": "Artificial Intelligence and Data Science"},
        {"code": "EC", "name": "Electronics and Communication Engineeering"},
        {"code": "ME", "name": "Mechanical Engineering"},
        {"code": "CE", "name": "Civil Engineering"},
        {"code": "IT", "name": "Information Technology"},
        {"code": "EE", "name": "Electrical & Electronics Engineering"}
    ]
    
    semesters = [{"value": i, "name": f"Semester {i}"} for i in range(1, 9)]
    
    return render_template('index.html', branches=branches, semesters=semesters)

@app.route('/process', methods=['POST'])
def process():
    first_roll = request.form.get('first_roll')
    last_roll = request.form.get('last_roll')
    semester = int(request.form.get('semester'))
    branch = request.form.get('branch')
    
    # Create a unique session ID for this job
    session_id = get_random_string()
    
    # Start processing in a background thread
    def background_task():
        processor = Processor(semester, branch)
        success = processor.start(first_roll, last_roll)
        
        if success:
            csv_data = processor.to_csv()
            if csv_data:
                # Store the CSV data in the app config for later download
                app.config[f'csv_{session_id}'] = csv_data
                socketio.emit('complete', {'session_id': session_id})
            else:
                socketio.emit('error', {'message': "Failed to generate CSV file."})
        else:
            socketio.emit('error', {'message': "Failed to process results."})
    
    # Run the task in a background thread
    threading.Thread(target=background_task).start()
    
    return jsonify({'session_id': session_id})

@app.route('/download/<session_id>')
def download(session_id):
    csv_data = app.config.get(f'csv_{session_id}')
    if not csv_data:
        return "CSV file not found", 404
    
    return Response(
        csv_data,
        mimetype="text/csv",
        headers={"Content-disposition": f"attachment; filename=results_{session_id}.csv"}
    )

if __name__ == "__main__":
    socketio.run(app, debug=True, allow_unsafe_werkzeug=True)