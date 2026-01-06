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

id = ['hsra@inzisoft.com','jkjeong@inzisoft.com','shpark@inzisoft.com', 'Min.jw@inzisoft.com', 'smjin@inzisoft.com', 'hwjung@inzisoft.com']                  # 테스트 계정 id
pw = ['1111','1111','1111', '1111', '1111', '1111']                                      # 테스트 계정 pw
lable = ['근태담당자','대표이사','기획', '경영지원', '팀장', '팀원']         # 권한별 계정 리스트에 담기

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

driver.find_element(By.ID, 'details-button').click()                    # 크롬 자세히 버튼 클릭
time.sleep(1)
driver.find_element(By.ID, 'proceed-link').click()                      # 크롬 계속하기 버튼 클릭
driver.implicitly_wait(3)

f = open("tbmslog.txt", "w")                                            # 메모장 생성 (exe 생성 후 콘솔창 대신 메모장에 로그 찍기 위해)
time.sleep (2)

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

    j = i + 1                                                               # 로그 넘버링 작업
    num = str(j) + ". " + lable[i] + " (" + id[i] + " )"
    print (num)
    f.write (num + "\r\n")

    try:                                                                      # 정상 로그인 불가할 경우를 위해 try 사용 
        driver.find_element(By.LINK_TEXT, '휴가').click()    # 휴가 탭 클릭
        time.sleep(3)
    except:
        error_msg = driver.find_element(By.XPATH, '//*[@id="errMsg"]').text     # 로그인 불가시, 에러 텍스트 가져오기
        print(error_msg)
        messagebox.showinfo("로그인 실패",error_msg)                             # 윈도우 메세지 띄우기
    else:
        try:
            vc_list = driver.find_element(By.LINK_TEXT, '직원 휴가 내역').text
            filter1 = driver.find_element(By.ID, 'companyArea').text
            filter2 = driver.find_element(By.ID, 'departmentArea').text
            filter3 = driver.find_element(By.ID, 'employeeArea').text
            print (vc_list + " : " + filter1 + ">" + filter2 + ">" + filter3)
            f.write (f"{vc_list} : {filter1} > {filter2} > {filter3} \r\n")
            time.sleep(1)

            vc_status = driver.find_element(By.LINK_TEXT, '직원 휴가 사용 현황').text
            driver.find_element(By.LINK_TEXT, '직원 휴가 사용 현황').click()    # 직원 휴가 사용 현황 탭 클릭
            time.sleep(1)
            filter1 = driver.find_element(By.ID, 'companyArea').text
            filter2 = driver.find_element(By.ID, 'departmentArea').text
            filter3 = driver.find_element(By.ID, 'employeeArea').text
            print (vc_status + " : " + filter1 + ">" + filter2 + ">" + filter3)
            f.write (f"{vc_status} : {filter1} > {filter2} > {filter3} \r\n")
        except :
            print ('직원 휴가 사용 현황 탭 미존재')
            print ()
            f.write ("직원 휴가 사용 현황 탭 미존재" + "\r\n")
            driver.find_element(By.ID, 'userNm').click()    # 퇴사자 휴가관리 탭 클릭
            time.sleep(1)
            driver.find_element(By.LINK_TEXT, '로그 아웃').click()    # 퇴사자 휴가관리 탭 클릭
            time.sleep(1)
            continue
        try:
            vc_monthlystatus = driver.find_element(By.LINK_TEXT, '월별 직원 휴가 내역').text
            driver.find_element(By.LINK_TEXT, '월별 직원 휴가 내역').click()    # 월별 직원 휴가 내역 탭 클릭
            time.sleep(1)
            filter1 = driver.find_element(By.ID, 'companyArea').text
            filter2 = driver.find_element(By.ID, 'departmentArea').text
            filter3 = driver.find_element(By.ID, 'employeeArea').text
            print (vc_monthlystatus + " : " + filter1 + ">" + filter2 + ">" + filter3)
            f.write (f"{vc_monthlystatus} : {filter1} > {filter2} > {filter3} \r\n")
        except :
            print ('월별 직원 휴가 내역 탭 미존재')
            print ()
            f.write ("월별 직원 휴가 내역 탭 미존재" + "\r\n")
            driver.find_element(By.ID, 'userNm').click()    # 퇴사자 휴가관리 탭 클릭
            time.sleep(1)
            driver.find_element(By.LINK_TEXT, '로그 아웃').click()    # 퇴사자 휴가관리 탭 클릭
            time.sleep(1)
            continue
        try:
            vc_resignmember = driver.find_element(By.LINK_TEXT, '퇴사자 휴가관리').text
            driver.find_element(By.LINK_TEXT, '퇴사자 휴가관리').click()    # 퇴사자 휴가관리 탭 클릭
            time.sleep(1)
            filter1 = driver.find_element(By.ID, 'companyArea').text
            print (vc_resignmember + " : " + filter1)
            f.write (f"{vc_resignmember} : {filter1} \r\n")
        except : 
            print ('퇴사자 휴가관리 탭 미존재')
            print ()
            f.write ("퇴사자 휴가관리 탭 미존재" + "\r\n")
            driver.find_element(By.ID, 'userNm').click()    # 퇴사자 휴가관리 탭 클릭
            time.sleep(1)
            driver.find_element(By.LINK_TEXT, '로그 아웃').click()    # 퇴사자 휴가관리 탭 클릭
            time.sleep(1)
            continue
        
    '''print (vc_status)
    print (vc_monthlystatus)
    print (vc_resignmember)'''

    print ("-" * 20)
    driver.find_element(By.ID, 'userNm').click()    # 퇴사자 휴가관리 탭 클릭
    time.sleep(1)
    driver.find_element(By.LINK_TEXT, '로그 아웃').click()    # 퇴사자 휴가관리 탭 클릭
    time.sleep(1)

f.close()
