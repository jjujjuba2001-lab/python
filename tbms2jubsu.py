import selenium
import random
import string
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
lable = ['근태담당자','대표이사','기획', '경영지원', '팀장', '팀원']         # 권한별 계정 리스트에 담기

driver = webdriver.Chrome()
Options = Options()
Options.add_experimental_option("detach", True)

def get_col_width_row_height(img_width, img_height):                    # 이미지 크기에 따라 셀 높이, 너비 조정 항수
    col_width = img_width*63.2/504.19
    row_height = img_height*225.35/298.96
    return (col_width, row_height)

driver.get("https://dev-tbms.mobileleader.com:8060/login")                 # TBMS2.0 사이트 접속
driver.maximize_window()
Options.add_experimental_option("excludeSwitches", ["enable-automation"])     # 자동화 제어 알림 줄 삭제 , 하지만 안됨
driver.implicitly_wait(5)
print (driver.current_window_handle)

path = r'C:\Users\user\Desktop\TBMS SCREEN SHOT'                                        # 저장할 스크린 샷 경로
root = tk.Tk()
root.withdraw()

driver.find_element(By.ID, 'details-button').click()                    # 크롬 자세히 버튼 클릭
time.sleep(1)
driver.find_element(By.ID, 'proceed-link').click()                      # 크롬 계속하기 버튼 클릭
driver.implicitly_wait(3)

'''f = open("tbmslog.txt", "w")                                            # 메모장 생성 (exe 생성 후 콘솔창 대신 메모장에 로그 찍기 위해)
time.sleep (2)'''

now = datetime.now()
formatted_now = now.strftime("%m-%d %H:%M")
title = " | " + str(formatted_now) + " " + "파이썬으로 제목 입력하기"           # 제목에 추가할 글자
text = "파이썬 글자 입력, this is test to verify the max length of input text. !@#$%^*()_+=李居銀1234567890,이거슨 100자!!" *10         # 1000자 만들기

for i in range(len(id)):                                                               # 계정별 반복 실행
    sdkid = driver.find_element(By.XPATH, '//*[@id="plainUserId"]')               
    sdkid.send_keys(id[i])                                                             # id 입력
    time.sleep(2)
    #sdkpw = driver.find_element(By.XPATH, '//*[@id="plainPassword"]')

    driver.find_element(By.XPATH, '//*[@id="plainPassword"]').send_keys(pw[i])           # pw 입력
    driver.implicitly_wait(5)

    #driver.find_element(By.XPATH, '//*[@id="loginForm"]/div[1]/div[1]/a').click()      # 로그인 버튼 클릭
    enter_btn = driver.find_element(By.LINK_TEXT, '로그인')
    enter_btn.send_keys(Keys.ENTER)                                                     # 로그인 버튼 엔터키 입력하기
    time.sleep(3)

    ''' j = i + 1                                                               # 로그 넘버링 작업
    num = str(j) + ". " + lable[i] + " (" + id[i] + " )"
    print (num)
    f.write (num + "\r\n")'''

    try:                                                                      # 정상 로그인 불가할 경우를 위해 try 사용 
        driver.find_element(By.LINK_TEXT, '기술지원').click()    # 휴가 탭 클릭
        time.sleep(3)
    except:
        alert = driver.switch_to.alert                          # 크롬 알럿창 제어를 위해
        print(alert.text)
        time.sleep(5)
        alert.accept()                                          # 크롬 알럿창 확인 클릭
        time.sleep(5)
        driver.find_element(By.LINK_TEXT, '기술지원').click()    # 기술지원 탭 클릭
        time.sleep(3)
    else:
        #driver.find_element(By.ID, 'isAllTxt').click()             # 전체 탭 선택
        driver.find_element(By.ID, 'isRequestTxt').click()         # 요청 탭 선택
        #driver.find_element(By.ID, 'isReceiptTxt').click()         # 접수 탭 선택
        #driver.find_element(By.ID, 'isWaitingTxt').click()         # 대기 탭 선택
        #driver.find_element(By.ID, 'isProcessingTxt').click()      # 진행 탭 선택   
        #driver.find_element(By.ID, 'isProcessingDelayTxt').click()  # 지연 탭 선택
        #driver.find_element(By.ID, 'isUnresolvedProcessingTxt').click()    # 미해결건 재진행 탭 선택  
        #driver.find_element(By.ID, 'isUnresolvedClosingTxt').click()       # 미해결 종료 탭 선택
        time.sleep(2)
        for k in range(5):
            char = string.ascii_letters + string.digits + string.punctuation
            rand_cus_staff = random.choices (char, k=3)
            rand_cus_staff = ''.join(rand_cus_staff)
            print (rand_cus_staff)
            rand_text = random.choices (char, k=200)
            rand_text = ''.join(rand_text)
            print (rand_text)

            driver.find_element (By.CLASS_NAME, 'ev_dhx_web').click()
            time.sleep(2)
            driver.find_element (By.ID, 'btn_receiptCustomerSupport').click()
            time.sleep(2)
            driver.find_element (By.ID, 'btn_openPopup').click()
            time.sleep(2)
            driver.find_element (By.ID, 'searchCustomer').click()
            time.sleep(2)
            driver.find_element (By.ID, 'searchCustomer').send_keys("ibk")
            time.sleep(2)
            driver.find_element (By.ID, '44').click()
            time.sleep(2)
            driver.find_element (By.ID, 'btn_selectCustomer').click()
            time.sleep(2)

            driver.find_element (By.ID, 'contractSelectBox').click()
            time.sleep(2)
            driver.find_element (By.ID, 'contractSelectBox').send_keys(Keys.DOWN)
            time.sleep(1)
            driver.find_element (By.ID, 'contractSelectBox').send_keys(Keys.ENTER)
            time.sleep(1)

            driver.find_element (By.ID, 'productSelectBox').click()
            time.sleep(2)
            driver.find_element (By.ID, 'productSelectBox').send_keys(Keys.DOWN)
            time.sleep(1)
            driver.find_element (By.ID, 'productSelectBox').send_keys(Keys.ENTER)
            time.sleep(1)

            driver.find_element (By.ID, 'requestSubject').click()
            time.sleep(2)
            driver.find_element (By.ID, 'requestSubject').send_keys(title)
            time.sleep(2)
            
            driver.find_element (By.ID, 'requestContent').click()
            time.sleep(2)
            driver.find_element (By.ID, 'requestContent').send_keys(Keys.ENTER)
            time.sleep(2)
            driver.find_element (By.ID, 'requestContent').send_keys(rand_text)
            time.sleep(2)

            driver.find_element (By.NAME, 'customerSupportTypeCode').click()
            time.sleep(2)
            driver.find_element (By.NAME, 'customerSupportTypeCode').send_keys(Keys.DOWN)
            time.sleep(1)
            driver.find_element (By.NAME, 'customerSupportTypeCode').send_keys(Keys.ENTER)
            time.sleep(1)

            driver.find_element (By.ID, 'customerStaffName').click()
            time.sleep(2)
            driver.find_element (By.ID, 'customerStaffName').send_keys(" | 파이썬 고객 " + str(rand_cus_staff))
            time.sleep(2)

            driver.find_element (By.ID, 'btn_modify').click()
            time.sleep(2)
        
            alert = driver.switch_to.alert                          # 크롬 알럿창 제어를 위해
            alert1 = alert.text
            alert_diff = "이메일을 정확히 입력해주세요."
            print(alert1)
            time.sleep(5)
            alert.accept()                                          # 크롬 알럿창 확인 클릭
            time.sleep(3)

            if alert_diff == alert1:
                driver.find_element(By.ID, 'customerStaffEmail').clear()    # 기술지원 탭 클릭
                driver.find_element(By.ID, 'customerStaffEmail').click()
                driver.find_element(By.ID, 'customerStaffEmail').send_keys("abc@python.com")
                time.sleep(1)
                driver.find_element (By.ID, 'btn_modify').click()
                time.sleep(2)

                alert = driver.switch_to.alert                          # 크롬 알럿창 제어를 위해
                print(alert.text)
                time.sleep(5)
                alert.accept()                                          # 크롬 알럿창 확인 클릭
                time.sleep(3)

                driver.find_element (By.LINK_TEXT, '메일발송').click()
                time.sleep(7)

                alert = driver.switch_to.alert                          # 크롬 알럿창 제어를 위해
                print(alert.text)
                time.sleep(5)
                alert.accept()                                          # 크롬 알럿창 확인 클릭
                time.sleep(3)
            else:
                driver.find_element (By.LINK_TEXT, '메일발송').click()
                time.sleep(7)

                alert = driver.switch_to.alert                          # 크롬 알럿창 제어를 위해
                print(alert.text)
                time.sleep(5)
                alert.accept()                                          # 크롬 알럿창 확인 클릭
                time.sleep(3)
                
            driver.find_element (By.LINK_TEXT, '기술지원').click()
            time.sleep(3)
            driver.find_element(By.ID, 'isRequestTxt').click()         # 요청 탭 선택
            time.sleep(3)

            '''full_scroll = driver.execute_script("return document.body.scrollHeight")
            while True:
                ev_list = driver.find_elements (By.CLASS_NAME, 'ev_dhx_web')
                odd_list = driver.find_elements (By.CLASS_NAME, 'odd_dhx_web')
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(1.5)
                scroll_height = driver.execute_script("return document.body.scrollHeight")
                print (scroll_height)
                if full_scroll == scroll_height:
                    break
                else:
                    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                    full_scroll = scroll_height
                    time.sleep(2)
        
        print (type(ev_list))
        print (len(ev_list))
        for ev in ev_list:
            print (ev.text)

        print (type(odd_list))
        print (len(odd_list))
        for odd in odd_list:
            print (odd.text)

        print ("-" * 20)
        print (int(len(ev_list)) + int(len(odd_list)))
        print ("-" * 20)
        driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(3)'''

    driver.find_element(By.CLASS_NAME, 'lnb_logout').click()    # 퇴사자 휴가관리 탭 클릭
    time.sleep(1)
    alert = driver.switch_to.alert  
    alert.accept()  