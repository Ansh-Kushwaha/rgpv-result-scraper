import os
import requests
import pytesseract
import random
import string
import tocsv
from io import BytesIO
from PIL import Image
from bs4 import BeautifulSoup

from concurrent.futures import ThreadPoolExecutor
from time import sleep
import threading

pytesseract.pytesseract.tesseract_cmd = 'C:/Program Files/Tesseract-OCR/tesseract'

def get_random_string():
    random_str = ''.join([random.choice(string.ascii_letters + string.digits) for _ in range(24)])
    return (random_str)

class Processor:
    fail = False
    processed_count = 0

    def __init__(self, sem):
        self.lock = threading.Lock()

        self.sem = sem

        self.first_entry=True
        self.worksheet=tocsv.tocsv()
        self.results = {}
        self.num_cols = False


    def start(self):
        sess_url = self.get_session()
        if self.fail:
            return()
        self.sess, self.url = sess_url
        self.process(wait = True)

    def process(self, wait = False):
        with ThreadPoolExecutor(max_workers = 200) as executor:
            for roll in self.roll_list_generator():
                executor.submit(self.try_open, roll)
            executor.shutdown(wait = wait)

    def try_open(self,roll):
        while self.get_result(roll) == 1:
            pass

    def get_session(self):
        try:
            cookie = get_random_string()
            header = {'User-Agent' : 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Mobile Safari/537.36 Edg/130.0.0.0','Cookies':'ASP.NET_SessionId=' + cookie}

            sess = requests.session()
            sess.headers.update(header)
            program_resp = sess.get('http://result.rgpv.ac.in/Result/ProgramSelect.aspx')

            soup = BeautifulSoup(program_resp.text,'html5lib')

            deptid = 'radlstProgram_1'
            value = soup.find('input',{'id':deptid})['value']
            deptid = deptid.replace('_','$')
            viewState = soup.find('input',{'id':'__VIEWSTATE'})['value']
            viewStateGen = soup.find('input',{'id':'__VIEWSTATEGENERATOR'})['value']
            EvenValidation = soup.find('input',{'id':'__EVENTVALIDATION'})['value']
            post_data = {'__EVENTTARGET':deptid,'__EVENTARGUMENT':'','__LASTFOCUS':'','__VIEWSTATE':viewState,'__VIEWSTATEGENERATOR':viewStateGen,'__EVENTVALIDATION':EvenValidation,'radlstProgram':value}
            resp=sess.post('http://result.rgpv.ac.in/Result/ProgramSelect.aspx',data=post_data,allow_redirects=True)
            url = resp.url
            return ((sess, url))

        except Exception as e:
            print("Exception while establishing session: ", e)
            self.fail = True

    def get_result(self, roll):
        for _ in range(10):
            try:
                with self.lock:
                    resp = self.sess.get(self.url)
                
                soup = BeautifulSoup(resp.text,'html5lib')
                image_url = "http://result.rgpv.ac.in/Result/" + soup.findAll('img')[1]['src']
                response = requests.get(image_url)
                
                if response.status_code != 200:
                    return (1)
                
                img = Image.open(BytesIO(response.content))
                solution = pytesseract.image_to_string(img).strip().upper().replace(' ', '')
                if not solution:
                    return(1)
                
                # change this time in case of exceptions (minimum 5)
                sleep(5)
                viewState = soup.find('input',{'id':'__VIEWSTATE'})['value']
                viewStateGen = soup.find('input',{'id':'__VIEWSTATEGENERATOR'})['value']
                EvenValidation = soup.find('input',{'id':'__EVENTVALIDATION'})['value']

                post_data = {'__EVENTTARGET':'', '__EVENTARGUMENT':'', '__LASTFOCUS':'', '__VIEWSTATE':viewState, '__VIEWSTATEGENERATOR':viewStateGen, '__EVENTVALIDATION':EvenValidation, 'ctl00$ContentPlaceHolder1$txtrollno':roll, 'ctl00$ContentPlaceHolder1$drpSemester':str(self.sem), 'ctl00$ContentPlaceHolder1$rbtnlstSType':'G', 'ctl00$ContentPlaceHolder1$TextBox1':solution, 'ctl00$ContentPlaceHolder1$btnviewresult':'View Result'}

                with self.lock:
                    result = self.sess.post(self.url, data=post_data, allow_redirects=True)

                result_found='<td class="resultheader">'
                wrong_captcha='<script language="JavaScript">alert("you have entered a wrong_captcha text");</script>'
                result_not_found='<script language=JavaScript>alert("Result for this Enrollment No. not Found");</script>'
                
                if result_found in result.text:
                    self.process_result(result.text, roll)
                    return(0)
                
                elif wrong_captcha in result.text:
                    return(1)
                elif result_not_found in result.text:
                    # todo handle this
                    return(0)
                else:
                    return(1)

            except Exception as e:
                print("Exception while opening result for roll: ", roll, e)
                self.fail = True
            else:
                break
        else:
            self.fail = True

    def process_result(self, html, roll):
        list = []
        soup = BeautifulSoup(html,'html5lib')

        name = soup.find(id="ctl00_ContentPlaceHolder1_lblNameGrading").get_text().strip()
        sgpa = soup.find(id="ctl00_ContentPlaceHolder1_lblSGPA").get_text()
        cgpa = soup.find(id="ctl00_ContentPlaceHolder1_lblcgpa").get_text()
        result = soup.find(id="ctl00_ContentPlaceHolder1_lblResultNewGrading").get_text()
        
        with self.lock:
            list.append(roll)
            list.append(name)
        
        results = soup.findAll("table")[0].findAll("table")[2].findAll("tr")[6].findAll("table")
        
        with self.lock:
            if self.first_entry is True:
                self.first_entry = False
                header_row = []
                header_row.append("Enrollment Number")
                header_row.append("Name")
                for row in range(1, len(results)):
                    header_row.append(results[row].findAll('td')[0].text.replace("\n",'').strip())
                header_row.append("SGPA")
                header_row.append("CGPA")
                header_row.append("Result")
                self.results[0] = header_row
                self.num_cols = len(header_row)


        for row in range(1, len(results)):
            list.append(results[row].findAll('td')[3].text.replace("\n",'').strip())

        list.append(sgpa)
        list.append(cgpa)
        list.append(result)
        with self.lock:
            self.processed_count += 1
            self.results[int(roll[-3:])] = list

            if self.processed_count % 10 == 0:
                print("Processed", self.processed_count, "students.")

    def to_csv(self):
        if self.fail:
            return(-1)
        
        list = sorted(self.results.items())
        if self.num_cols:
            self.worksheet.fromlist(list)
            return(self.worksheet.getcsv())
        else:
            return(-1)


    def roll_list_generator(self): 
        f = input('Enter First Enrollment number: ')
        l = input('Enter Last Enrollment number: ')

        if len(f) != len(l):
            print("Incorrect enrollment numbers.")
            return ()

        roll_list = list()
        start = int(f[-4:])
        end = int(l[-4:]) + 1
        common = f[:8]
        for i in range(start,end):
            i = str(i)
            roll = common + i
            roll_list.append(roll)
        
        return(roll_list)

    

if __name__ == "__main__":
    semester = int(input("Enter Semester (1 to 8): "))
    filename = input("Enter filename: ")
    res_processor = Processor(semester)
    res_processor.start()
    resp = res_processor.to_csv()
    assert(resp not in [-1]), resp

    if filename[-4:] != ".csv":
        filename = filename + ".csv"
    
    counter = 1
    base = filename[:-4]
    while os.path.exists(filename):
        filename = f"{base}({counter}).csv"
        counter += 1
    with open(filename,'w+') as file:
        file.write(resp)