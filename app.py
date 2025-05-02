import os
import requests
import pytesseract
import random
import string
import logging
from io import BytesIO, StringIO
from PIL import Image
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor
from time import sleep
import threading
import csv
from datetime import datetime

from flask import Flask, render_template, request, jsonify, send_file, Response, url_for
from flask_socketio import SocketIO

app = Flask(__name__)
socketio = SocketIO(app)

# Setup logging
def setup_logger():
    # Get the directory where the script is located
    script_dir = os.path.dirname(os.path.abspath(__file__))
    logs_dir = os.path.join(script_dir, 'logs')
    
    # Create logs directory if it doesn't exist
    if not os.path.exists(logs_dir):
        os.makedirs(logs_dir)
    
    # Generate timestamp for log file name with local time
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_file = os.path.join(logs_dir, f'rgpv_scraper_{timestamp}.log')
    
    # Configure logger
    logger = logging.getLogger('rgpv_scraper')
    logger.setLevel(logging.INFO)
    
    # File handler
    file_handler = logging.FileHandler(log_file)
    file_handler.setLevel(logging.INFO)
    
    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    
    # Create formatter
    formatter = logging.Formatter('%(asctime)s | %(levelname)-8s | %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)
    
    # Add handlers to logger
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger

# Initialize logger
logger = setup_logger()

# Configure Tesseract path
if os.name == 'nt':  # Windows
    pytesseract.pytesseract.tesseract_cmd = 'C:/Program Files/Tesseract-OCR/tesseract'
    logger.info("Configured Tesseract for Windows environment")
else:  # Linux/macOS
    pytesseract.pytesseract.tesseract_cmd = '/usr/bin/tesseract'
    logger.info("Configured Tesseract for Linux/macOS environment")

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
        logger.info(f"Initialized Processor for branch: {branch}, semester: {sem}")

    def start(self, first_roll, last_roll):
        logger.info(f"Starting process for rolls {first_roll} to {last_roll}")
        self.roll_list = self.roll_list_generator(first_roll, last_roll)
        self.total_count = len(self.roll_list)
        logger.info(f"Generated roll list with {self.total_count} enrollment numbers")
        
        logger.info("Establishing session with RGPV server...")
        sess_url = self.get_session()
        if self.fail:
            logger.error("Failed to establish session")
            return False
        self.sess, self.url = sess_url
        logger.info(f"Session established successfully. URL: {self.url}")
        
        logger.info("Starting processing of roll numbers")
        self.process(wait=True)
        logger.info("Processing completed")
        return True

    def process(self, wait=False):
        logger.info(f"Creating thread pool with max_workers=50")
        with ThreadPoolExecutor(max_workers=50) as executor:
            for roll in self.roll_list:
                executor.submit(self.try_open, roll)
            logger.info("All tasks submitted to thread pool")
            executor.shutdown(wait=wait)
            logger.info("Thread pool shutdown complete")

    def try_open(self, roll):
        logger.debug(f"Attempting to process roll: {roll}")
        attempts = 0
        while self.get_result(roll) == 1:
            attempts += 1
            logger.debug(f"Retry #{attempts} for roll: {roll}")
            pass

    def get_session(self):
        try:
            logger.info("Generating session cookie and headers")
            cookie = get_random_string()
            header = {
                'User-Agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Mobile Safari/537.36 Edg/130.0.0.0',
                'Cookies': 'ASP.NET_SessionId=' + cookie
            }

            logger.info("Creating new session with RGPV server")
            sess = requests.session()
            sess.headers.update(header)
            
            logger.info("Requesting program selection page")
            program_resp = sess.get('http://result.rgpv.ac.in/Result/ProgramSelect.aspx')
            logger.info(f"Program selection page status code: {program_resp.status_code}")

            logger.info("Parsing response with BeautifulSoup")
            soup = BeautifulSoup(program_resp.text, 'html5lib')

            deptid = 'radlstProgram_1'
            value = soup.find('input', {'id': deptid})['value']
            deptid = deptid.replace('_', '$')
            viewState = soup.find('input', {'id': '__VIEWSTATE'})['value']
            viewStateGen = soup.find('input', {'id': '__VIEWSTATEGENERATOR'})['value']
            EvenValidation = soup.find('input', {'id': '__EVENTVALIDATION'})['value']
            logger.info("Extracted form values successfully")
            
            post_data = {
                '__EVENTTARGET': deptid,
                '__EVENTARGUMENT': '',
                '__LASTFOCUS': '',
                '__VIEWSTATE': viewState,
                '__VIEWSTATEGENERATOR': viewStateGen,
                '__EVENTVALIDATION': EvenValidation,
                'radlstProgram': value
            }
            
            logger.info("Submitting program selection form")
            resp = sess.post('http://result.rgpv.ac.in/Result/ProgramSelect.aspx', data=post_data, allow_redirects=True)
            url = resp.url
            logger.info(f"Redirected to URL: {url}")
            
            return (sess, url)

        except Exception as e:
            socketio.emit('error', {'message': f"Session error: {str(e)}"})
            self.fail = True

    def get_result(self, roll):
        for attempt in range(10):
            try:
                logger.debug(f"Attempt #{attempt+1} for roll: {roll}")
                with self.lock:
                    resp = self.sess.get(self.url)
                
                soup = BeautifulSoup(resp.text, 'html5lib')
                image_url = "http://result.rgpv.ac.in/Result/" + soup.findAll('img')[1]['src']
                logger.debug(f"CAPTCHA image URL: {image_url}")
                
                response = requests.get(image_url)
                
                if response.status_code != 200:
                    logger.warning(f"Failed to fetch CAPTCHA image (status: {response.status_code})")
                    return 1
                
                logger.debug("Processing CAPTCHA image with Tesseract OCR")
                img = Image.open(BytesIO(response.content))
                solution = pytesseract.image_to_string(img, config='--psm 7 --oem 1').strip().upper().replace(' ', '')
                if not solution:
                    logger.warning("Empty CAPTCHA solution returned by Tesseract")
                    return 1
                
                logger.debug(f"CAPTCHA solution: {solution}")
                
                # minimum 5 second delay
                logger.debug("Waiting 5 seconds before submitting form (rate limiting)")
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

                logger.debug(f"Submitting form for roll: {roll}, semester: {self.sem}")
                with self.lock:
                    result = self.sess.post(self.url, data=post_data, allow_redirects=True)

                result_found = '<td class="resultheader">'
                wrong_captcha = '<script language="JavaScript">alert("you have entered a wrong_captcha text");</script>'
                result_not_found = '<script language=JavaScript>alert("Result for this Enrollment No. not Found");</script>'
                
                if result_found in result.text:
                    logger.info(f"Result found for roll: {roll}")
                    self.process_result(result.text, roll)
                    return 0
                
                elif wrong_captcha in result.text:
                    logger.warning(f"Wrong CAPTCHA for roll: {roll}, will retry")
                    return 1
                elif result_not_found in result.text:
                    logger.info(f"No result found for roll: {roll}")
                    with self.lock:
                        self.processed_count += 1
                        progress = int((self.processed_count / self.total_count) * 100)
                        socketio.start_background_task(
                            socketio.emit, 'progress', 
                            {'count': self.processed_count, 'total': self.total_count, 'percent': progress}
                        )
                    logger.debug(f"Progress: {self.processed_count}/{self.total_count} ({progress}%)")
                    return 0
                else:
                    logger.warning(f"Unknown response for roll: {roll}, will retry")
                    return 1

            except Exception as e:
                socketio.emit('error', {'message': f"Error processing {roll}: {str(e)}"})
                self.fail = True
            else:
                break
        else:
            logger.error(f"Failed after 10 attempts for roll: {roll}")
            self.fail = True

    def process_result(self, html, roll):
        logger.debug(f"Processing result HTML for roll: {roll}")
        list_data = []
        soup = BeautifulSoup(html, 'html5lib')

        name = soup.find(id="ctl00_ContentPlaceHolder1_lblNameGrading").get_text().strip()
        sgpa = soup.find(id="ctl00_ContentPlaceHolder1_lblSGPA").get_text()
        cgpa = soup.find(id="ctl00_ContentPlaceHolder1_lblcgpa").get_text()
        result = soup.find(id="ctl00_ContentPlaceHolder1_lblResultNewGrading").get_text()
        
        logger.debug(f"Student: {name}, Roll: {roll}, SGPA: {sgpa}, CGPA: {cgpa}, Result: {result}")
        
        with self.lock:
            list_data.append(roll)
            list_data.append(name)
        
        results = soup.findAll("table")[0].findAll("table")[2].findAll("tr")[6].findAll("table")
        
        with self.lock:
            if self.first_entry is True:
                logger.info("Processing first entry to create header row")
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
                logger.info(f"Created header row with {self.num_cols} columns")

        for row in range(1, len(results)):
            list_data.append(results[row].findAll('td')[3].text.replace("\n", '').strip())

        list_data.append(sgpa)
        list_data.append(cgpa)
        list_data.append(result)
        
        with self.lock:
            self.processed_count += 1
            self.results[int(roll[-3:])] = list_data
            
            progress = int((self.processed_count / self.total_count) * 100)
            socketio.start_background_task(
                socketio.emit, 'progress', 
                {'count': self.processed_count, 'total': self.total_count, 'percent': progress}
            )
            logger.debug(f"Progress: {self.processed_count}/{self.total_count} ({progress}%)")


    def to_csv(self):
        logger.info("Converting results to CSV format")
        if self.fail:
            logger.error("Cannot generate CSV due to previous failures")
            return None
        
        sorted_list = sorted(self.results.items())
        if self.num_cols:
            # Convert to CSV string
            output = StringIO()
            writer = csv.writer(output)
            for _, row in sorted_list:
                writer.writerow(row)
            logger.info(f"CSV generated successfully with {len(sorted_list)} rows")
            return output.getvalue()
        else:
            logger.warning("No columns defined, cannot generate CSV")
            return None

    def roll_list_generator(self, first_roll, last_roll):
        logger.info(f"Generating roll list from {first_roll} to {last_roll}")
        if len(first_roll) != len(last_roll):
            logger.error("Enrollment numbers have different lengths")
            socketio.emit('error', {'message': "Incorrect enrollment numbers format."})
            return []

        roll_list = []
        start = int(first_roll[-4:])
        end = int(last_roll[-4:]) + 1
        common = first_roll[:8]
        logger.debug(f"Common prefix: {common}, range: {start}-{end-1}")
        
        for i in range(start, end):
            i = str(i).zfill(4)  # Zero-pad to ensure consistent length
            roll = common + i
            roll_list.append(roll)
        
        logger.info(f"Generated {len(roll_list)} enrollment numbers")
        return roll_list

# Flask routes
@app.route('/')
def index():
    logger.info("Serving index page")
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
    
    logger.info(f"Processing request for branch: {branch}, semester: {semester}")
    logger.info(f"Roll range: {first_roll} to {last_roll}")
    
    # Create a unique session ID for this job
    session_id = get_random_string()
    logger.info(f"Created session ID: {session_id}")
    
    # Start processing in a background thread
    def background_task():
        logger.info(f"Starting background task for session: {session_id}")
        processor = Processor(semester, branch)
        success = processor.start(first_roll, last_roll)
        
        if success:
            logger.info("Processing completed successfully")
            csv_data = processor.to_csv()
            if csv_data:
                # Store the CSV data in the app config for later download
                app.config[f'csv_{session_id}'] = csv_data
                logger.info(f"CSV data stored for session: {session_id}")
                socketio.emit('complete', {'session_id': session_id})
            else:
                logger.error("Failed to generate CSV file")
                socketio.emit('error', {'message': "Failed to generate CSV file."})
        else:
            logger.error("Failed to process results")
            socketio.emit('error', {'message': "Failed to process results."})
    
    # Run the task in a background thread
    logger.info("Launching background thread")
    threading.Thread(target=background_task).start()
    
    return jsonify({'session_id': session_id})

@app.route('/download/<session_id>')
def download(session_id):
    logger.info(f"Download request for session: {session_id}")
    csv_data = app.config.get(f'csv_{session_id}')
    if not csv_data:
        logger.error(f"CSV data not found for session: {session_id}")
        return "CSV file not found", 404
    
    logger.info(f"Sending CSV file for session: {session_id}")
    return Response(
        csv_data,
        mimetype="text/csv",
        headers={"Content-disposition": f"attachment; filename=results_{session_id}.csv"}
    )

port = int(os.environ.get("PORT", 5000))
logger.info(f"Starting server on port {port}")
socketio = SocketIO(app, cors_allowed_origins="*", async_mode="threading", ping_timeout=60, ping_interval=25)
logger.info("Socket.IO initialized with threading mode")
logger.info("Starting application... Press Ctrl+C to exit")
socketio.run(app, host="0.0.0.0", port=port, debug=True, allow_unsafe_werkzeug=True)