from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from openpyxl import Workbook
from openpyxl.drawing.image import Image

current_folder = 'D:/my_project/selim/'

options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ["enable-logging"])  
options.add_argument('--start-maximized')  
options.add_argument('--ignore-certificate-errors')
options.add_argument('--ignore-ssl-errors')
#options.add_argument('headless')    # 크롬을 백그라운드에서 실행
options.add_argument( "user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_6) \
    AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36")     # 봇으로 인식하지 않게끔 설정

browser = webdriver.Chrome(options=options)
browser.implicitly_wait(10) # seconds

#전기요금
url1 = 'https://www.billkorea.co.kr/ebill/index.ac'
browser.get(url1)
time.sleep(1)

#ID/PWD 입력하고 로그인
browser.find_element(By.XPATH, '//*[@id="loginForm"]/table/tbody/tr/td[1]/input').send_keys('01033181821')
browser.find_element(By.XPATH, '//*[@id="loginForm"]/table/tbody/tr/td[3]/input').send_keys('soon11&bk')
browser.find_element(By.XPATH, '//*[@id="loginForm"]/table/tbody/tr/td[5]/a').click()
time.sleep(1)

browser.find_element(By.XPATH, 
'//*[@id="menu_main"]/table/tbody/tr[2]/td/table/tbody/tr/td[4]/table/tbody/tr[5]/td/table/tbody/tr[1]/td/a').click()
time.sleep(1)

browser.find_element(By.XPATH, '//*[@id="table1"]/tbody/tr[1]/td[10]/a').click()
time.sleep(1)

#윈도우와 프레임 전환
browser.switch_to.window(browser.window_handles[-1])
time.sleep(1)
element = browser.find_elements(By.ID, 'bill_f')
browser.switch_to.frame(element[-1])

e_price = []    #[기본요금,전력량요금,기후환경요금,연료비조정액,역률요금,인터넷빌링할인,전력기금,청구요금]
p = []
p.append((browser.find_element(By.XPATH, '/html/body/div[3]/table/tbody/tr[1]/td[1]/table[1]/tbody/tr[1]/td')).text)    #기본요금
p.append((browser.find_element(By.XPATH, '/html/body/div[3]/table/tbody/tr[1]/td[1]/table[1]/tbody/tr[2]/td')).text)    #전력량요금
p.append((browser.find_element(By.XPATH, '/html/body/div[3]/table/tbody/tr[1]/td[1]/table[1]/tbody/tr[3]/td')).text)    #기후환경요금
p.append((browser.find_element(By.XPATH, '/html/body/div[3]/table/tbody/tr[1]/td[1]/table[1]/tbody/tr[4]/td')).text)    #연료비조정액
p.append((browser.find_element(By.XPATH, '/html/body/div[3]/table/tbody/tr[1]/td[1]/table[1]/tbody/tr[5]/td')).text)    #지상역률료
#p.append((browser.find_element(By.XPATH, '/html/body/div[3]/table/tbody/tr[1]/td[1]/table[1]/tbody/tr[7]/td')).text)   #인터넷빌링 할인
p.append((browser.find_element(By.XPATH, '/html/body/div[3]/table/tbody/tr[1]/td[1]/table[1]/tbody/tr[9]/td')).text)    #전력기금
p.append((browser.find_element(By.XPATH, '/html/body/div[3]/table/tbody/tr[1]/td[1]/table[1]/tbody/tr[13]/td')).text)   #청구요금

for i in p:
    e_price.append(int(i.replace(',','').replace('원','')))
u = e_price[1]+e_price[2]+e_price[3]    #사용요금 합계
e_price.append(u)
#print(e_price)

browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
time.sleep(1)
element1 = browser.find_element(By.XPATH,'/html/body/div[3]')
element1_png = element1.screenshot_as_png
filename1 = current_folder + '전기요금.png'
with open(filename1, "wb") as file:
    file.write(element1_png)
time.sleep(1)

#수도요금
browser = webdriver.Chrome(options=options)
browser.implicitly_wait(10) # seconds
url2 = 'https://water.suwon.go.kr/waterpay/ncoe/index.do' #수원시 상수도사업소 사이버요금창구
browser.get(url2)
time.sleep(1)

# 수용가번호 입력 
browser.find_element(By.ID, 'fsuyNo1').send_keys('1014')
browser.find_element(By.ID, 'fsuyNo2').send_keys('101')
browser.find_element(By.ID, 'fsuyNo3').send_keys('118')
browser.find_element(By.ID, 'fsuyNo4').send_keys('1300')
browser.find_element(By.ID, 'fsuyNo5').send_keys('00')
browser.find_element(By.CLASS_NAME, 'myButton.ac').click()

w_price = []    #[상수도요금,하수도요금,물이용부담금,요금합계]
p = []
p.append((browser.find_element(By.XPATH, '//*[@id="trList"]/td[5]')).text)       
p.append((browser.find_element(By.XPATH, '//*[@id="trList"]/td[6]')).text)         
p.append((browser.find_element(By.XPATH, '//*[@id="trList"]/td[8]')).text)         
p.append((browser.find_element(By.XPATH, '//*[@id="trList"]/td[10]')).text)     

for i in p:
    w_price.append(int(i.replace(',','')))
#print(w_price)

element2 = browser.find_element(By.XPATH,'//*[@id="Sub_Cont_content"]')
element2_png = element2.screenshot_as_png
filename2 = current_folder + '수도요금.png'
with open(filename2, "wb") as file:
    file.write(element2_png)
time.sleep(1)

#이미지 생성
img1 = Image(filename1)
img1.height = 200
img1.width = 320

img2 = Image(filename2) 
img2.height = 200
img2.width = 320

#엑셀파일 "자동조회"에 전기/수도요금 저장
filename = current_folder + '자동조회.xlsx'
wb = Workbook()
sheet = wb.active

sheet.append(['기본요금','전력량요금','기후환경요금','연료비조정액','역률요금','전력기금','청구요금','사용요금'])
sheet.append(e_price)
sheet.append(['상수도요금','하수도요금','물이용부담금','수도요금합계'])
sheet.append(w_price)
sheet.insert_rows(3)
sheet.add_image(img1, 'A8')
sheet.add_image(img2, 'F8')
wb.save(filename)


'''
#엑셀파일 "AutoTest" (임대료내역 파일형태)에 직접 전기/수도요금 저장
filename = current_folder + 'AutoTest.xlsx'
wb = openpyxl.load_workbook(filename)
sheet = wb.get_sheet_by_name('전기수도요금')

sheet['E4'].value = e_price[0]
sheet['F4'].value = e_price[1]+e_price[2]+e_price[3]
sheet['G4'].value = -(e_price[4]+e_price[5]+e_price[6])
sheet['C5'].value = e_price[7]
sheet['C25'].value = w_price[0]
sheet['C26'].value = w_price[1]
sheet['C27'].value = w_price[2]

pic = sheet._images[:]
sheet._images.remove(pic[0])
sheet._images.remove(pic[1])

sheet.add_image(img1, 'A41')
sheet.add_image(img2, 'E41')

wb.save(filename)
'''