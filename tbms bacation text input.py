import selenium
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.options import Options

from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait

import tkinter as tk                    # 윈도우 에러 메시지 띄우기 위해 import
from tkinter import messagebox

import openpyxl                                 # 엑셀 컨트롤
from openpyxl.styles import Border
from openpyxl.drawing.image import Image            # 엑셀에 이미지 삽입 위해
from PIL import Image as PILImage

id = ['hsra@inzisoft.com']                  # 테스트 계정 id
pw = ['1111']                                      # 테스트 계정 pw
lable = ['기획,경영지원,근태관리(라한솔)']         # 권한별 계정 리스트에 담기

driver = webdriver.Chrome()
Options = Options()
Options.add_experimental_option("detach", True)

def get_col_width_row_height(img_width, img_height):                    # 이미지 크기에 따라 셀 높이, 너비 조정 항수
    col_width = img_width*63.2/504.19
    row_height = img_height*225.35/298.96
    return (col_width, row_height)

driver.get("https://dev-tbms.mobileleader.com:8050/login")                 # TBMS 사이트 접속
driver.maximize_window()
Options.add_experimental_option("excludeSwitches", ["enable-automation"])     # 자동화 제어 알림 줄 삭제 , 하지만 안됨
driver.implicitly_wait(5)
print (driver.current_window_handle)

path = r'C:\Users\user\Desktop\TBMS SCREEN SHOT'                                        # 저장할 스크린 샷 경로
root = tk.Tk()
root.withdraw()

driver.find_element(By.ID, 'details-button').click()
time.sleep(1)
driver.find_element(By.ID, 'proceed-link').click()
driver.implicitly_wait(3)

for i in range(len(id)):                                                               # 계정별 반복 실행
    sdkid = driver.find_element(By.XPATH, '//*[@id="plainUserId"]')               
    sdkid.send_keys(id[i])                                                             # id 입력
    time.sleep(2)
    #sdkpw = driver.find_element(By.XPATH, '//*[@id="plainUserPw"]')

    driver.find_element(By.XPATH, '//*[@id="plainUserPw"]').send_keys(pw[i])           # pw 입력
    driver.implicitly_wait(5)

    #driver.find_element(By.XPATH, '//*[@id="loginForm"]/div[1]/div[1]/a').click()      # 로그인 버튼 클릭
    enter_btn = driver.find_element(By.XPATH, '//*[@id="loginForm"]/div[1]/div[1]/a')
    enter_btn.send_keys(Keys.ENTER)                                                     # 로그인 버튼 엔터키 입력하기
    time.sleep(3)

    try:                                                                      # 정상 로그인 불가할 경우를 위해 try 사용 
        driver.find_element(By.LINK_TEXT, '휴가').click()    # 휴가 탭 클릭
        time.sleep(3)
    except:
        error_msg = driver.find_element(By.XPATH, '//*[@id="errMsg"]').text     # 로그인 불가시, 에러 텍스트 가져오기
        print(error_msg)
        messagebox.showinfo("로그인 실패",error_msg)                             # 윈도우 메세지 띄우기
    else:
        driver.find_element(By.LINK_TEXT, '등록하기').click()    # 등록하기 클릭
        time.sleep(1)

        '''driver.find_element(By.XPATH, '//*[@id="stffTags"]/li/input').send_keys("아무개 프로")       #제목 입력하기
        time.sleep(1)
        driver.find_element(By.XPATH, '//*[@id="stffTags"]/li/input').send_keys(Keys.ENTER)'''

        '''iframe = driver.find_element(By.CSS_SELECTOR, 'iframe')                 # 본문 입력을 위해 iframe 찾기
        driver.switch_to.frame(iframe)                                          # 본문 입력을 위해 iframe 진입'''
        driver.find_element(By.ID, 'startDatePopup').click()                    # 시작일 인풋박스 클릭
        time.sleep(1)

        '''iframe = driver.find_element(By.CSS_SELECTOR, 'iframe')                 # 본문 입력을 위해 iframe 찾기 (iframe 총 2개라 1번더 찾기)
        driver.switch_to.frame(iframe)                                          # 본문 입력을 위해 iframe 진입 (iframe 총 2개라 1번더 진입)'''
        driver.find_element(By.CLASS_NAME, 'ui-state-default.ui-state-highlight').click()       # 데이트 피커에서 오늘 날짜 클릭
        #driver.find_element(By., '27').click()
        time.sleep(1)

        driver.find_element(By.ID, 'endDatePopup').click()                    # 종료일 인풋박스 클릭
        time.sleep(1)

        driver.find_element(By.CLASS_NAME, 'ui-state-default.ui-state-highlight').click()       # 데이트 피커에서 오늘 날짜 클릭
        time.sleep(1)
        
        driver.find_element(By.ID, 'vacTypeCdSelect').click()                       # 콤보 박스 클릭
        time.sleep(1)

        text1 = driver.find_element(By.XPATH, '//*[@id="vacTypeCdSelect"]/option[3]')           # 콤보 박스에서 휴가 구분 선택
        print(text1.text)
        driver.find_element(By.XPATH, '//*[@id="vacTypeCdSelect"]/option[3]').click()          
        time.sleep(1)

        driver.find_element(By.ID, 'vacTypeCdSelect').click()                        # 콤보 박스 다시 클릭
        time.sleep(1)
        
        now = datetime.now()                                            
        formatted_now = now.strftime("%m-%d %H:%M")
        title = (str(formatted_now) + " " + "파이썬으로" + text1.text + "입력하기") * 5

        element = driver.find_element(By.ID, 'vacRsn')             # 본문 입력을 위치 엘레멘트 찾기
        #text = "파이썬 글자 입력, this is test to verify the max length of input text. !@#$%^*()_+=李居銀1234567890,이거슨 100자!!" *20         # 2000자 만들기
        print(len(title))                                        # 글자 수 확인하기
        element.send_keys(title)                                 # 글자 입력
        time.sleep(1)
        #driver.switch_to.default_content()                      # iframe 벗어나 메인으로 돌아가기
        #time.sleep(1)
        driver.find_element(By.LINK_TEXT, '확인').click()       # 확인 버튼 클릭
        time.sleep(1)

        alert = driver.switch_to.alert                          # 크롬 알럿창 제어를 위해
        print(alert.text)
        alert.accept()                                          # 크롬 알럿창 확인 클릭
        
        #driver.find_element(By.LINK_TEXT, title).click()



        #driver.execute_script("arguments[0].scrollIntoView(true);", element)
        #element.click()
        #txt_area = driver.find_element(By.TAG_NAME, 'textarea')
        #time.sleep(2)
        #p_element = element.find_element(By.XPATH, "p")
        #text = "파이썬 글자 입력, this is test to verify the max length of input text. !@#$%^*()_+=李居銀1234567890,이거슨 100자!!" *20
        #element.send_keys(text)
        #driver.execute_script("arguments[0].innerHTML = '<span>{}</span>';".format(text), element)         # 자바 스크립트 이용한 강제 글씨입력, (html 구조 변경 가능성 있어서 비추천)
        #driver.switch_to.default_content() 

        '''for i in range (21):
            text = str(i) + "번째 글자 입력, this is test to verify the max length of input text. !@#$%^*()_+=李居銀1234567890,이거슨 100자!!"
            print(len(text))
            if i == 20:
                driver.execute_script("arguments[0].innerHTML = '<P>{}</P>';".format("끝"), element)
                print(len(text))
            else:  
                driver.execute_script("arguments[0].innerHTML = '<P>{}</P>';".format(text*20), element)
        #element.find_element(By.XPATH, "p").send_keys (str(i) + (text * 10) + "끝")'''

        
        # ==================================================================================== 에러메시지 제어 추가 해야함

        
