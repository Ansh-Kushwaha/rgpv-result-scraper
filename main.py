import os
import requests
import pytesseract
import time
import xlsxwriter
from PIL import Image
from bs4 import BeautifulSoup

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

pytesseract.pytesseract.tesseract_cmd = 'C:/Program Files/Tesseract-OCR/tesseract'

# global variables
driver = webdriver.Chrome()
wb = None
sheet = None
row_idx = 0
sheet_init = False
subs = []

# downloading and 
def clear_captcha():
    global driver
    l = list()
    captcha_elem = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_TextBox1")
    captcha_elem.clear()

    images = driver.find_elements(By.TAG_NAME, 'img')
    for image in images:
        a = image.get_attribute('src')
        l.append(a)
    captcha_src = l[1]
    response = requests.get(captcha_src)

    if response.status_code == 200:
        with open("sample.jpg", 'wb') as f:
            f.write(response.content)
    img = Image.open('sample.jpg')
    text = pytesseract.image_to_string(img).strip().upper().replace(' ', '')
    captcha_elem.send_keys(text)
    time.sleep(4.5)

# trying to open the result
def open_result(roll, num_sub, semester):
    global driver
    # base url
    driver.get("http://result.rgpv.ac.in/Result")
    
    btech_elem = driver.find_element(By.ID, "radlstProgram_0")
    btech_elem.send_keys(Keys.ARROW_RIGHT)
    roll_elem = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtrollno")
    roll_elem.send_keys(roll)
    sem_elem = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_drpSemester")
    for _ in range(1, semester):
        sem_elem.send_keys(Keys.DOWN)
    
    try:
        clear_captcha()
        driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnviewresult").send_keys(Keys.ENTER)

        wait = WebDriverWait(driver, timeout=0.5)
        alert_present = wait.until(EC.alert_is_present(), message="")

        if alert_present:
            alert = driver.switch_to.alert
            text = alert.text
            alert.accept()

            if "not Found" in text:
                print("Result not found for: ", roll)

                with open("./not_found.txt", "a") as f:
                    f.write(roll + "\n")
                pass
            elif "wrong text" in text:
                open_result(roll, num_sub, semester)
    except TimeoutException:
        store_result(roll, num_sub, semester)

# store result after opening
def store_result(roll, num_subs, sem):
    global driver

    resultSheet = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_pnlGrading")
    html = resultSheet.get_attribute("innerHTML")

    with open("./successful.txt", "a") as f:
        f.write(roll + "\n")
    
    soup = BeautifulSoup(html, 'lxml')
    name = soup.find(id="ctl00_ContentPlaceHolder1_lblNameGrading").get_text().strip()
    sgpa = soup.find(id="ctl00_ContentPlaceHolder1_lblSGPA").get_text()
    cgpa = soup.find(id="ctl00_ContentPlaceHolder1_lblcgpa").get_text()
    result = soup.find(id="ctl00_ContentPlaceHolder1_lblResultNewGrading").get_text()
    grades = []   
    base = 15

    global subs
    for i in range(0, num_subs):
        if len(subs) != num_subs:
            subs.append(soup.find_all('td')[base].get_text())
        grades.append(soup.find_all('td')[base + 3].get_text())
        base += 4

    global sheet_init
    if sheet_init == False:
        init_sheet(roll[4:6], sem)
        sheet_init = True

    global row_idx
    sheet.write(row_idx, 0, int(row_idx))
    sheet.write(row_idx, 1, name)
    sheet.write(row_idx, 2, roll)
    i = 3
    for grade in grades:
        sheet.write(row_idx, i, grade)
        i += 1
    sheet.write(row_idx, i, float(sgpa))
    i += 1
    sheet.write(row_idx, i, float(cgpa))
    i += 1
    sheet.write(row_idx, i, result)
    i += 1
    row_idx += 1

# set the first row of the sheet
def init_sheet(branch: str, sem):
    new_sheet = wb.add_worksheet(branch + '-Sem' + str(sem))
    bold = wb.add_format({'bold': True})
    new_sheet.write(0, 0, 'S. No', bold)
    new_sheet.write(0, 1, 'Name', bold)
    new_sheet.write(0, 2, 'Enrollment', bold)
    i = 3
    for sub in subs:
        new_sheet.write(0, i, sub, bold)
        i += 1
    new_sheet.write(0, i, 'SGPA', bold)
    i += 1
    new_sheet.write(0, i, 'CGPA', bold)
    i += 1
    new_sheet.write(0, i, 'Result', bold)
    i += 1
    global sheet
    sheet = new_sheet
    global row_idx
    row_idx += 1

# generate a list of roll numbers between a range
def roll_list_generator(): 
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

def create_workbook(branch_code, semester):
    global wb
    base = './Result-' + branch_code + "-" + str(semester) 
    new_filename = base + ".xlsx"
    counter = 1
    while os.path.exists(new_filename):
        new_filename = f"{base}({counter}).xlsx"
        counter += 1
    
    wb = xlsxwriter.Workbook(new_filename)

def main():
    num_sub = int(input("Enter number of subjects (check from result the number of rows in the table): "))
    semester = int(input("Enter Semester (1 to 8): "))
    branch_code = input("Enter branch code (Ex.: CS, EC, AD, AL, etc): ")

    with open("./successful.txt", "w") as f:
        f.write(branch_code + "-" + str(semester) + "Sem\n")
    with open("./not_found.txt", "w") as f:
        f.write(branch_code + "-" + str(semester) + "Sem\n")

    create_workbook(branch_code, semester)
    try:
        for roll in roll_list_generator():
            open_result(roll, num_sub, semester)
    except:
        sheet.autofit()


main()
if wb:
    wb.close()
driver.close()