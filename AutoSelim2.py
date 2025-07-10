import time
from datetime import datetime
import openpyxl
from openpyxl.drawing.image import Image
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import chromedriver_autoinstaller

# --- 설정 파일 임포트 ---
import config

# --- 상수 설정 (엑셀 구조) ---
# 데이터가 업데이트될 시트의 인덱스 (0부터 시작)
DATA_SHEET_INDEX = 5 
# 이미지가 업데이트될 시트들의 인덱스 범위
IMAGE_SHEET_INDICES = range(5, 12)

# 전기요금 데이터가 들어갈 셀 주소
E_CELLS = {
    '기본요금': 'E4',
    '사용요금합계': 'F4',
    '역률요금': 'G4', # 음수(-)로 들어감
    '청구요금': 'C5'
}

# 수도요금 데이터가 들어갈 셀 주소
W_CELLS = {
    '상수도요금': 'C25',
    '하수도요금': 'C26',
    '물이용부담금': 'C27'
}

# 이미지가 삽입될 셀 주소
IMAGE_CELLS = {
    '전기': 'A43',
    '수도': 'E43'
}

def setup_driver():
    """Chrome 드라이버를 설정하고 초기화합니다."""
    print("Chrome 드라이버를 설정하고 초기화합니다...")
    try:
        # chromedriver_autoinstaller를 사용하여 chromedriver를 설치하거나 업데이트합니다.
        chromedriver_autoinstaller.install()

        options = webdriver.ChromeOptions()
        options.add_experimental_option("excludeSwitches", ["enable-logging"])
        options.add_argument('--start-maximized')
        options.add_argument('--ignore-certificate-errors')
        options.add_argument('--ignore-ssl-errors')
        options.add_argument("user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36")

        # Service 객체 없이 바로 webdriver.Chrome을 호출합니다.
        # chromedriver_autoinstaller가 chromedriver의 경로를 자동으로 설정해줍니다.
        driver = webdriver.Chrome(options=options)
        print("드라이버 설정 완료.")
        return driver
    except Exception as e:
        print(f"드라이버 설정 중 오류가 발생했습니다: {e}")
        return None

def get_bill_data(driver, wait):
    """전기 및 수도 요금 정보를 조회하고 반환합니다."""
    # 1. 전기요금 조회
    print("전기요금 조회를 시작합니다...")
    driver.get('https://www.billkorea.co.kr/ebill/index.ac')
    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="loginForm"]/table/tbody/tr/td[1]/input'))).send_keys(config.BILLKOREA_ID)
    driver.find_element(By.XPATH, '//*[@id="loginForm"]/table/tbody/tr/td[3]/input').send_keys(config.BILLKOREA_PW)
    driver.find_element(By.XPATH, '//*[@id="loginForm"]/table/tbody/tr/td[5]/a').click()
    
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="menu_main"]/table/tbody/tr[2]/td/table/tbody/tr/td[4]/table/tbody/tr[5]/td/table/tbody/tr[1]/td/a'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="table1"]/tbody/tr[1]/td[10]/a'))).click()
    
    wait.until(EC.number_of_windows_to_be(2))
    driver.switch_to.window(driver.window_handles[1])
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'bill_f')))
    
    price_xpaths = [
        '/html/body/div[3]/table/tbody/tr[1]/td[1]/table[1]/tbody/tr[1]/td', # 기본요금
        '/html/body/div[3]/table/tbody/tr[1]/td[1]/table[1]/tbody/tr[2]/td', # 전력량요금
        '/html/body/div[3]/table/tbody/tr[1]/td[1]/table[1]/tbody/tr[3]/td', # 기후환경요금
        '/html/body/div[3]/table/tbody/tr[1]/td[1]/table[1]/tbody/tr[4]/td', # 연료비조정액
        '/html/body/div[3]/table/tbody/tr[1]/td[1]/table[1]/tbody/tr[5]/td', # 지상역률료
        '/html/body/div[3]/table/tbody/tr[1]/td[1]/table[1]/tbody/tr[12]/td' # 청구요금
    ]
    e_price_raw = [int(wait.until(EC.presence_of_element_located((By.XPATH, xp))).text.replace(',', '').replace('원', '')) for xp in price_xpaths]
    e_price = {
        '기본요금': e_price_raw[0],
        '사용요금합계': e_price_raw[1] + e_price_raw[2] + e_price_raw[3],
        '역률요금': e_price_raw[4],
        '청구요금': e_price_raw[5]
    }
    e_screenshot_path = config.WORKING_DIRECTORY + '전기요금.png'
    wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[3]'))).screenshot(e_screenshot_path)
    print("전기요금 정보 추출 완료.")
    driver.close()
    driver.switch_to.window(driver.window_handles[0])

    # 2. 수도요금 조회
    print("수도요금 조회를 시작합니다...")
    driver.get('https://water.suwon.go.kr/waterpay/ncoe/index.do')
    for i, num in enumerate(config.WATER_SUWON_NO, 1):
        driver.find_element(By.ID, f'fsuyNo{i}').send_keys(num)
    driver.find_element(By.CLASS_NAME, 'myButton.ac').click()
    
    w_price_xpaths = {
        '상수도요금': '//*[@id="trList"]/td[5]',
        '하수도요금': '//*[@id="trList"]/td[6]',
        '물이용부담금': '//*[@id="trList"]/td[8]'
    }
    w_price = {key: int(wait.until(EC.presence_of_element_located((By.XPATH, xp))).text.replace(',', '')) for key, xp in w_price_xpaths.items()}
    w_screenshot_path = config.WORKING_DIRECTORY + '수도요금.png'
    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="Sub_Cont_content"]'))).screenshot(w_screenshot_path)
    print("수도요금 정보 추출 완료.")

    return e_price, w_price, e_screenshot_path, w_screenshot_path

def update_monthly_report(e_price, w_price, e_img_path, w_img_path):
    """현재 달의 엑셀 보고서를 열어 데이터를 업데이트합니다."""
    # 파일명 자동 생성 (예: 2025-07.xlsx)
    filename = datetime.now().strftime("%Y-%m") + ".xlsx"
    filepath = config.MONTHLY_REPORT_DIRECTORY + filename
    print(f"월별 보고서 '{filepath}' 업데이트를 시작합니다...")

    try:
        wb = openpyxl.load_workbook(filepath)
        ws_list = wb.sheetnames

        # 1. 데이터 시트 업데이트
        data_sheet = wb[ws_list[DATA_SHEET_INDEX]]
        data_sheet[E_CELLS['기본요금']].value = e_price['기본요금']
        data_sheet[E_CELLS['사용요금합계']].value = e_price['사용요금합계']
        data_sheet[E_CELLS['역률요금']].value = -e_price['역률요금'] # 음수로 입력
        data_sheet[E_CELLS['청구요금']].value = e_price['청구요금']
        
        data_sheet[W_CELLS['상수도요금']].value = w_price['상수도요금']
        data_sheet[W_CELLS['하수도요금']].value = w_price['하수도요금']
        data_sheet[W_CELLS['물이용부담금']].value = w_price['물이용부담금']
        print("데이터 시트 업데이트 완료.")

        # 2. 이미지 시트 업데이트 (효율적으로 처리)
        img1 = Image(e_img_path)
        img1.height = 200
        img1.width = 320

        img2 = Image(w_img_path)
        img2.height = 200
        img2.width = 320

        for i in IMAGE_SHEET_INDICES:
            sheet = wb[ws_list[i]]
            # 기존 이미지 모두 제거 (더 안정적인 방법)
            sheet._images = []
            # 새 이미지 추가
            sheet.add_image(img1, IMAGE_CELLS['전기'])
            sheet.add_image(img2, IMAGE_CELLS['수도'])
        print("이미지 시트 업데이트 완료.")

        # 3. 파일 한 번만 저장
        wb.save(filepath)
        print(f"성공적으로 '{filepath}' 파일을 저장했습니다.")

    except FileNotFoundError:
        print(f"[오류] 파일을 찾을 수 없습니다: {filepath}")
    except Exception as e:
        print(f"[오류] 엑셀 파일 처리 중 문제가 발생했습니다: {e}")

def main():
    """메인 실행 함수"""
    driver = None
    try:
        driver = setup_driver()
        if driver is None:
            # 드라이버 설정에 실패하면 프로그램을 종료합니다.
            print("\n*** 드라이버 설정 실패. 프로그램을 종료합니다. ***")
            return

        wait = WebDriverWait(driver, 10)
        e_price, w_price, e_img, w_img = get_bill_data(driver, wait)
        update_monthly_report(e_price, w_price, e_img, w_img)
        print("\n*** 모든 작업이 성공적으로 완료되었습니다. ***")
    except Exception as e:
        print(f"\n*** 작업 중 오류가 발생했습니다: {e} ***")
    finally:
        if driver:
            driver.quit()
            print("브라우저를 종료했습니다.")

if __name__ == "__main__":
    main()
