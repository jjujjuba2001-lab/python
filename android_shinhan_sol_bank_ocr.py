from appium import webdriver
import selenium
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.actions.action_builder import ActionBuilder
from selenium.webdriver.common.actions.pointer_input import PointerInput
from selenium.webdriver.common.keys import Keys
from appium.webdriver.common.appiumby import AppiumBy
from appium.options.common import AppiumOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver import ActionChains
from selenium.webdriver.support import expected_conditions as EC
import subprocess
from time import sleep

# AppiumOptions 객체 생성
options = AppiumOptions()
options.set_capability("platformName", "Android")
options.set_capability("platformVersion", "13")                 # 안드로이드 버전
options.set_capability("deviceName", "R54TC00EBPY")
options.set_capability("appPackage", "")                        # 초기 실행할 앱 설정 가능    
options.set_capability("appActivity", "")                       # 초기 실행할 앱의 액티비티 입력  
options.set_capability("automationName", "UiAutomator2")
options.set_capability("autoGrantPermissions", "true")          # 권한 자동 승인 , 하지만 동작 안됨


CName = ['정현우']
CBirth = [830831]
CBirth_last = ["1234567"]

# Appium 서버와 연결
driver = webdriver.Remote('http://127.0.0.1:4723', options=options)

# 앱이 실행될 때까지 잠시 대기
sleep(2)

# 설치된 앱 목록 가져오기

cmd = "adb shell pm list packages | findstr sbanking"                                  # cmd 에 넣을 명령어를 변수에 대입 (앱 패키지 리스트 중 sbanking 으로 필터)
# cwd= 로 특정 경로로 이동후 cmd 명령어 입력
app_package_name = subprocess.run(cmd, capture_output=True, text=True, shell=True, cwd= r"C:\Users\mung9\Desktop\adb\platform-tools")

if app_package_name.returncode == 0:
    print("ADB 출력 결과:")
    print(app_package_name.stdout)
else:
    print("오류 발생:")
    print(app_package_name.stderr)

app_package_name = app_package_name.stdout
temp = app_package_name.split ("package:")                  # package: 기준으로 자르기
app_package_name = temp[1]                                  # package: 기준 뒷 부분만 사용 하기
print (app_package_name)

driver.activate_app (app_package_name)                      # 앱 실행
driver.implicitly_wait (10)

# 권한 팝업이 나타날 때까지 기다리기
try:
    allow_button = WebDriverWait(driver, 3).until(
        EC.element_to_be_clickable((AppiumBy.ID, 'com.android.permissioncontroller:id/permission_allow_foreground_only_button'))     # 앱 사용중에만 허용 버튼
    )
    allow_button.click()                            # 앱 사용중에만 허용 버튼 클릭

except Exception as e:
    print(f"버튼을 찾는 중 오류 발생: {e}")                 # 권한 팝업 미 발생시 에러 발생

finally:
    driver.find_element (AppiumBy.ID, 'com.shinhan.sbanking:id/button').click()                    # '신한 sol 뱅크 시작하기' 버튼 클릭
    WebDriverWait(driver, 10)
    driver.find_element (AppiumBy.XPATH, '//android.widget.Button[@resource-id="atomicButton5"]').click()                     # '인증서나 신분증으로 본인인증을 할 수 있어요' 버튼 클릭
    WebDriverWait(driver, 15)                                                                                                    # 초 대기
    driver.find_element (AppiumBy.XPATH, '//android.widget.CheckBox[@resource-id="checkboxUID190-label"]').click()                     # '[필수] 개인정보 수집·이용 동의서 (모바일채널 본인확인)' 버튼 클릭
    WebDriverWait(driver, 10)

    driver.find_element (AppiumBy.XPATH, '//android.widget.EditText[@resource-id="customerName"]').send_keys(CName[0])                  # 고객명 입력 필드에 고객명 입력
    WebDriverWait(driver, 5)

    elements = driver.find_elements (AppiumBy.CLASS_NAME, 'android.widget.EditText')                                    # EditText 클래스 수집

    for i in elements:
        print (i)

    # 고객 생년월일 입력 필드에 
    CBirth_input = driver.find_element (AppiumBy.ANDROID_UIAUTOMATOR, 'new UiSelector().text("주민등록번호 앞 6자리, 6자리 입력 시 주민번호 7번째 입력 서식으로 이동합니다")')
    #driver.execute_script ("mobile: hideKeyboard")                     # 가상 키보드 숨기기
    WebDriverWait(driver, 10)
    
    driver.find_element (AppiumBy.XPATH, '//android.widget.EditText[@text="주민등록번호 앞 6자리, 6자리 입력 시 주민번호 7번째 입력 서식으로 이동합니다"]').click()
    WebDriverWait(driver, 10)
    driver.execute_script("mobile: shell", {
    "command": f"input text {CBirth[0]}"
    })

    #WebDriverWait(driver, 10)
    #driver.find_element (AppiumBy.XPATH, '//android.widget.LinearLayout[@resource-id="com.shinhan.sbanking:id/numberSKeypad"]').click()
    WebDriverWait(driver, 20)
    
    for num in CBirth_last[0]:
        try:
            # `numberSKeypad` 요소가 존재하는지 확인
            print (num)
            numskey = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((AppiumBy.XPATH, '//android.widget.LinearLayout[@resource-id="com.shinhan.sbanking:id/numberSKeypad"]'))
            )
            # `numskey` 요소가 찾았으므로, 버튼 클릭
            WebDriverWait(driver, 10)
            driver.find_element(AppiumBy.XPATH, f"//android.widget.LinearLayout[@content-desc='{num} 버튼']").click()
        except Exception as e:
            print(f"❌ 오류 발생: {e}")
            break  # 요소를 찾을 수 없으면 반복문 종료

    driver.find_element (AppiumBy.XPATH, '//android.widget.EditText[@resource-id="customerPhoneNumber"]').click()
    WebDriverWait(driver, 10)
    elements[2].send_keys('01012345678')

    driver.find_element (AppiumBy.XPATH, '//android.widget.Button[@resource-id="atomicButton13"]').click()                     # 시작 버튼 클릭
    WebDriverWait(driver, 10)                                                                                                    # 3초 대기
    try:
        allow_button = WebDriverWait(driver, 3).until(
        EC.element_to_be_clickable((AppiumBy.XPATH, '//android.widget.Button[@resource-id="com.android.permissioncontroller:id/permission_allow_foreground_only_button"]'))     # 앱 사용중에만 허용 버튼
    )
        allow_button.click()                            # 앱 사용중에만 허용 버튼 클릭

    except Exception as e:
        print(f"버튼을 찾는 중 오류 발생: {e}")                 # 권한 팝업 미 발생시 에러 발생

    WebDriverWait(driver, 20)
    driver.find_element (AppiumBy.XPATH, '//android.widget.Button[@resource-id="check"]').click()                     # 촬영 화면 진입후 촬영 버튼 클릭 
    WebDriverWait(driver, 20)

    driver.find_element (AppiumBy.ACCESSIBILITY_ID, '주민등록증, 운전면허증').click()        # 카메라 모드 콤보박스 클릭
    WebDriverWait(driver, 20)   

    try:
        result_screen = WebDriverWait(driver, 30).until(
        EC._element_if_visible((AppiumBy.XPATH, '//android.widget.TextView[@text="인식된 정보를 확인해주세요"]'))     # 앱 사용중에만 허용 버튼
    )
                                  # 앱 사용중에만 허용 버튼 클릭

    except Exception as e:
        print(f"버튼을 찾는 중 오류 발생: {e}")                 # 권한 팝업 미 발생시 에러 발생
 




