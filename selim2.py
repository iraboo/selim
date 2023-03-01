from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains

import chromedriver_autoinstaller
import pyautogui
import time
import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl import Workbook
from openpyxl.drawing.image import Image



current_folder = 'C:/Users/iraboo/Documents/my_project/selim/'
chrome_ver = chromedriver_autoinstaller.get_chrome_version().split('.')[0]  #크롬드라이버 버전 확인
chrome_file = current_folder + chrome_ver + '/chromedriver'

options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ["enable-logging"])    
options.add_argument('--ignore-certificate-errors')
options.add_argument('--ignore-ssl-errors')
#options.add_argument('headless')    # 크롬을 백그라운드에서 실행
options.add_argument( "user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_6) \
    AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36")     # 봇으로 인식하지 않게끔 설정
ser = Service(chrome_file)

try:
    browser = webdriver.Chrome(service=ser, options=options)
except:     # 크롬버전이 다르면 ./{chrome_ver}에 다시 설치
    chromedriver_autoinstaller.install(True)
    browser = webdriver.Chrome(service=ser, options=options)

#browser = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options, executable_path='C:/Users/iraboo/Documents/my_project/KiwoomTrading/test/chromedriver')
browser = webdriver.Chrome(service=ser, options=options)
browser.implicitly_wait(10) # seconds

#전기요금
url1 = 'https://home.kepco.co.kr/kepco/main.do'
browser.get(url1)
browser.maximize_window()
time.sleep(1)
login_bt = browser.find_element(By.CLASS_NAME, 'login')
login_bt.click()
time.sleep(1)

id = browser.find_element(By.XPATH, '//*[@id="id_A"]')
id.send_keys('swkwak11')
pwd = browser.find_element(By.XPATH, '//*[@id="pw_A"]')
pwd.send_keys('soon11')
login1_bt = browser.find_element(By.XPATH, '//*[@id="content"]/div/div[1]/div/form/div/div/div/div[4]')
login1_bt.click()
time.sleep(1)

next_bt = browser.find_element(By.XPATH, '//*[@id="goLogin"]')
next_bt.click()
time.sleep(1)

cyber_bt = browser.find_element(By.XPATH, '//*[@id="nav"]/ul/li[3]/a')
cyber_bt.click()
time.sleep(1)

query_bt = browser.find_element(By.XPATH, '//*[@id="content"]/div/div[8]/div[1]/a[2]')
#query_bt.click()
query_bt.send_keys(Keys.ENTER)
time.sleep(1)

#팝업창 처리 필요
#//*[@id="wrapPop"]/dl/dd[2]/a
pop_bt = browser.find_element(By.XPATH, '//*[@id="wrapPop"]/dl/dd[2]')
pop_bt.click()
time.sleep(1)

query1_bt = browser.find_element(By.XPATH, '//*[@id="content"]/div[5]/div[1]/table/tbody/tr[1]')
#query1_bt = browser.find_element(By.XPATH, '//*[@id="content"]/div[6]/div[1]/table/tbody/tr[1]')
#query1_bt.click()
query1_bt.click()

#기본요금
price11 = (browser.find_element(By.XPATH, 
                               '//*[@id="content"]/div[3]/div[1]/div[1]/div/table/tbody/tr[1]/td/div')).text
#전력량요금
price12 = (browser.find_element(By.XPATH, 
                               '//*[@id="content"]/div[3]/div[1]/div[1]/div/table/tbody/tr[2]/td/div')).text
#기후환경요금
price13 = (browser.find_element(By.XPATH, 
                               '//*[@id="content"]/div[3]/div[1]/div[1]/div/table/tbody/tr[3]/td/div')).text
#연료비조정액
price14 = (browser.find_element(By.XPATH, 
                               '//*[@id="content"]/div[3]/div[1]/div[1]/div/table/tbody/tr[4]/td/div')).text
#역률요금
price15 = (browser.find_element(By.XPATH, 
                               '//*[@id="content"]/div[3]/div[1]/div[1]/div/table/tbody/tr[5]/td/div')).text
#자동납부할인할인
price16 = (browser.find_element(By.XPATH, 
                               '//*[@id="content"]/div[3]/div[1]/div[1]/div/table/tbody/tr[6]/td/div')).text
#인터넷할인
price17 = (browser.find_element(By.XPATH, 
                               '//*[@id="content"]/div[3]/div[1]/div[1]/div/table/tbody/tr[7]/td/div')).text
#전력기금
price18 = (browser.find_element(By.XPATH, 
                               '//*[@id="content"]/div[3]/div[1]/div[1]/div/table/tbody/tr[10]/td/div')).text



element = browser.find_element(By.XPATH,'//*[@id="content"]/div[3]/div[1]/div[2]')
action = ActionChains(browser)
action.move_to_element(element).perform()

element1 = browser.find_element(By.XPATH,'//*[@id="content"]/div[3]/div[1]/div[1]')

element1_png = element1.screenshot_as_png

filename1 = current_folder + '전기요금.png'
with open(filename1, "wb") as file:
    file.write(element1_png)

time.sleep(1)
browser.close()
time.sleep(3)

#수도요금
browser = webdriver.Chrome(service=ser, options=options)
browser.implicitly_wait(10) # seconds
url2 = 'https://water.suwon.go.kr/waterpay/ncoe/index.do' #수원시 상수도사업소 사이버요금창구
browser.get(url2)
browser.maximize_window()
time.sleep(1)

# 수용가번호 입력 
#id=browser.find_element_by_id('fsuyNo1')
id = browser.find_element(By.ID, 'fsuyNo1')
id.send_keys('1014')

id = browser.find_element(By.ID, 'fsuyNo2')
id.send_keys('101')

id = browser.find_element(By.ID, 'fsuyNo3')
id.send_keys('118')

id = browser.find_element(By.ID, 'fsuyNo4')
id.send_keys('1300')

id = browser.find_element(By.ID, 'fsuyNo5')
id.send_keys('00')

query_bt = browser.find_element(By.CLASS_NAME, 'myButton.ac')
query_bt.click()

price21 = (browser.find_element(By.XPATH, '//*[@id="trList"]/td[5]')).text           #상수도요금
price22 = (browser.find_element(By.XPATH, '//*[@id="trList"]/td[6]')).text           #하수도요금
price23 = (browser.find_element(By.XPATH, '//*[@id="trList"]/td[8]')).text           #물이용부담금
due_date = (browser.find_element(By.XPATH, '//*[@id="trList"]/td[9]')).text         #납기일
price24 = (browser.find_element(By.XPATH, '//*[@id="trList"]/td[10]')).text       #요금합계

filename2 = current_folder + '수도요금.png'
#element = browser.find_element(By.CLASS_NAME,'area_left')
element = browser.find_element(By.XPATH,'//*[@id="Sub_Cont_content"]')

element_png = element.screenshot_as_png
with open(filename2, "wb") as file:
    file.write(element_png)
    
time.sleep(1)
browser.close()


#엑셀에 전기/수도요금 저장
filename = current_folder + 'AutoTest.xlsx'
wb = openpyxl.load_workbook(filename)
sheet = wb.get_sheet_by_name('전기수도요금')

'''
sheet.append(['기본요금','전력량요금','기후환경요금','연료비조정액','역률요금','자동납부할인','인터넷할인','전력기금'])
sheet.append([int(price11.replace(',','')),int(price12.replace(',','')),int(price13.replace(',','')),int(price14.replace(',','')),int(price15.replace(',','')),int(price16.replace(',','')),int(price17.replace(',','')),int(price18.replace(',',''))])

sheet.append(['상수도요금','하수도요금','물이용부담금','수도요금합계'])
sheet.append([int(price21.replace(',','')),int(price22.replace(',','')),int(price23.replace(',','')),int(price24.replace(',',''))])

sheet.insert_rows(3)
'''

p1 = int(price11.replace(',',''))                                                                   #기본요금
p2 = int(price12.replace(',',''))+int(price13.replace(',',''))+int(price14.replace(',',''))         #사용요금
p3 = -(int(price15.replace(',',''))+int(price16.replace(',',''))+int(price17.replace(',','')))      #할인요금
p4 = int(price18.replace(',',''))                                                                   #전력기금

p5 = int(price21.replace(',',''))                                                                   #상수도요금
p6 = int(price22.replace(',',''))                                                                   #하수도요금
p7 = int(price23.replace(',',''))                                                                   #물이용분담금

sheet['E4'].value = p1
sheet['F4'].value = p2
sheet['G4'].value = p3
sheet['C5'].value = p4
sheet['C25'].value = p5
sheet['C26'].value = p6
sheet['C27'].value = p7

img1 = Image(filename1) 
img1.height = 200
img1.width = 310

img2 = Image(filename2) 
img2.height = 200
img2.width = 310

pic = sheet._images[:]
sheet._images.remove(pic[0])
sheet._images.remove(pic[1])

sheet.add_image(img1, 'A41')
sheet.add_image(img2, 'E41')

wb.save(filename)
