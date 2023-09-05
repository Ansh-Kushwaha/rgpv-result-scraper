import requests
import pytesseract
import time

from PIL import Image
from bs4 import BeautifulSoup
from xlwt import Workbook
from xlrd import open_workbook
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.alert import Alert

pytesseract.pytesseract.tesseract_cmd = 'C:/Program Files/Tesseract-OCR/tesseract'
driver = webdriver.Chrome()
wb = Workbook()
subs = []
sheetInit = False
sheet = None
rowNum = 0


def scrape(roll, subNum, sem, driver):
    driver.get("http://result.rgpv.ac.in/Result")
    openPage(roll, subNum, sem, driver) 

def openPage(roll, subNum, sem, driver):
    btechElem = driver.find_element(By.ID, "radlstProgram_0")
    btechElem.send_keys(Keys.ARROW_RIGHT)
    inputData(roll, subNum, sem, driver)

def inputData(roll, subNum, sem, driver):
    rollElem = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtrollno")
    rollElem.send_keys(roll)
    semElem = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_drpSemester")
    for _ in range(1, sem):
        semElem.send_keys(Keys.DOWN)
    
    clearCaptcha(roll, subNum, sem, driver)

def clearCaptcha(roll, subNum, sem, driver):
    l = list()
    captchaElem = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_TextBox1")
    captchaElem.clear()

    images = driver.find_elements(By.TAG_NAME, 'img')
    for image in images:
        a = image.get_attribute('src')
        l.append(a)
    captchaSrc = l[1]
    response = requests.get(captchaSrc)
    if response.status_code == 200:
        with open("sample.jpg", 'wb') as f:
            f.write(response.content)
    img = Image.open('sample.jpg')
    text = pytesseract.image_to_string(img).strip().upper().replace(' ', '')
    print(text)
    captchaElem.send_keys(text)
    time.sleep(5)
    openResult(roll, subNum, sem, driver)

def openResult(roll, subNum, sem, driver):
    driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnviewresult").send_keys(Keys.ENTER)
    try:
        Alert(driver).accept()
        clearCaptcha(roll, subNum, sem, driver)
    except:
        storeResult(roll, subNum, sem, driver)


def storeResult(roll, subNum, sem, driver):
    resultSheet = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_pnlGrading")
    html = resultSheet.get_attribute("innerHTML")

    with open("parsed.txt", "a") as f:
        f.write(roll + "\n")
    parse(html, subNum, sem, roll)

def parse(html, subNum, sem, roll):
    soup = BeautifulSoup(html, 'lxml')

    name = soup.find(id="ctl00_ContentPlaceHolder1_lblNameGrading").get_text().strip()
    sgpa = soup.find(id="ctl00_ContentPlaceHolder1_lblSGPA").get_text()
    cgpa = soup.find(id="ctl00_ContentPlaceHolder1_lblcgpa").get_text()
    result = soup.find(id="ctl00_ContentPlaceHolder1_lblResultNewGrading").get_text()
    grades = []   
    base = 15
    for i in range(0, subNum):
        if len(subs) != subNum:
            subs.append(soup.find_all('td')[base].get_text())
        grades.append(soup.find_all('td')[base + 3].get_text())
        base += 4

    global sheetInit
    if sheetInit == False:
        initSpreadsheet(roll[5:7], sem)
        sheetInit = True
    enterResult(name, roll, grades, sgpa, cgpa, result)
    # print(name, sgpa, cgpa, result, grades)

def enterResult(name, roll, grades, sgpa, cgpa, result):
    global rowNum
    sheet.write(rowNum, 0, int(rowNum))
    sheet.write(rowNum, 1, name)
    sheet.write(rowNum, 2, roll)
    i = 3
    for grade in grades:
        sheet.write(rowNum, i, grade)
        i += 1
    sheet.write(rowNum, i, float(sgpa))
    i += 1
    sheet.write(rowNum, i, float(cgpa))
    i += 1
    sheet.write(rowNum, i, result)
    i += 1
    rowNum += 1

def initSpreadsheet(branch: str, sem):
    sheet1 = wb.add_sheet(branch + '-Sem' + str(sem))
    sheet1.write(0, 0, 'S. No')
    sheet1.write(0, 1, 'Enrollment')
    sheet1.write(0, 2, 'Name')
    # print(subs)
    i = 3
    for sub in subs:
        sheet1.write(0, i, sub)
        i += 1
    sheet1.write(0, i, 'SGPA')
    i += 1
    sheet1.write(0, i, 'CGPA')
    i += 1
    sheet1.write(0, i, 'Result')
    i += 1
    global sheet
    sheet = sheet1
    global rowNum
    rowNum += 1
    # wb.save('Result.xls')

def rollListGen(): 
    f = input('Enter First Roll number: ')
    l = input('Enter Last Roll number: ')
    roll_list = list()
    start = int(f[-4:])
    end = int(l[-4:]) + 1
    common = f[:8]
    for i in range(start,end):
        i = str(i)
        roll = common + i
        roll_list.append(roll)
    return(roll_list)


def main():
    subNum = int(input("Enter number of subjects (check from result upto last): "))
    sem = int(input("Enter Semester: "))
    for roll in rollListGen():
        scrape(roll, subNum, sem, driver)
    wb.save('Result.xls')

main()
driver.close()