import requests
import pytesseract
import time

from PIL import Image
from bs4 import BeautifulSoup
from xlwt import Workbook
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select

pytesseract.pytesseract.tesseract_cmd = 'C:/Program Files/Tesseract-OCR/tesseract'
driver = webdriver.Chrome()
wb = Workbook()

def inputData(roll, sem, driver):
    rollElem = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtrollno")
    rollElem.send_keys(roll)
    semElem = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_drpSemester")
    for _ in range(1, sem):
        semElem.send_keys(Keys.DOWN)
    
    clearCaptcha(driver)

def clearCaptcha(driver):
    l = list()
    captchaElem = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_TextBox1")
    # captchaElem.clear()

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
    time.sleep(5)
    print(text)
    try:
        captchaElem.send_keys(text)
        captchaElem.send_keys(Keys.ENTER)
    except:
        pass

def openPage(roll, sem, driver):
    btechElem = driver.find_element(By.ID, "radlstProgram_0")
    btechElem.send_keys(Keys.ARROW_RIGHT)
    inputData(roll, sem, driver)

def scrape(roll, sem, driver):
    driver.get("http://result.rgpv.ac.in/Result")
    openPage(roll, sem, driver)

def initSpreadsheet(branch: str, sem):

    sheet1 = wb.add_sheet(branch + 'Sem' + str(sem))
    sheet1.write(0, 0, 'S. No')
    sheet1.write(0, 1, 'Enrollment')
    sheet1.write(0, 2, 'Name')
    sheet1.write(0, 3, 'Sub1')
    sheet1.write(0, 4, 'Sub2')
    sheet1.write(0, 5, 'Sub3')
    sheet1.write(0, 6, 'Sub4')
    sheet1.write(0, 7, 'Sub5')
    sheet1.write(0, 8, 'SGPA')
    sheet1.write(0, 9, 'CSPA')

    wb.save('Result.xls')

def rollListGen(): 
    #first = input('Enter First Roll number: ')
    #last = input('Enter Last Roll number: ')
    f = '0103CS211001'
    l = '0103CS211073'
    roll_list = list()
    start = int(f[-4:])
    end = int(l[-4:])+1
    common = f[:8]
    for i in range(start,end):
        i = str(i)
        roll = common + i
        roll_list.append(roll)
    return(roll_list)


def main():
    for roll in rollListGen():
        parsed = open('parsed.txt','r+').read().split()
        if roll not in parsed:
            scrape(roll, 4, driver)
            print('parsing for ', roll)

main()
driver.close()
# initSpreadsheet('CSE', 4)
# scrape("0103CS211030", 4, driver)