from appium import webdriver
import selenium
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.actions.action_builder import ActionBuilder
from selenium.webdriver.common.actions.pointer_input import PointerInput
from appium.webdriver.common.appiumby import AppiumBy
from appium.options.common import AppiumOptions
from datetime import datetime
from selenium.webdriver.support.ui import WebDriverWait
from time import sleep

# AppiumOptions 객체 생성
options = AppiumOptions()
options.set_capability("platformName", "Android")
options.set_capability("platformVersion", "15")  # 안드로이드 버전
options.set_capability("deviceName", "R54W601GV4X")
options.set_capability("appPackage", "")
options.set_capability("appActivity", "")
options.set_capability("automationName", "UiAutomator2")

# Appium 서버와 연결
driver = webdriver.Remote('http://127.0.0.1:4723', options=options)

# 앱이 실행될 때까지 잠시 대기
sleep(2)

# 계산기 앱 실행후
driver.find_element(AppiumBy.ACCESSIBILITY_ID, "7").click()
driver.find_element(AppiumBy.ID, "com.sec.android.app.popupcalculator:id/calc_keypad_btn_add").click()
driver.find_element(AppiumBy.ID, "com.sec.android.app.popupcalculator:id/calc_keypad_btn_03").click()

# 계산 결과 확인 (예: "=" 버튼 클릭 후)
driver.find_element(AppiumBy.ID, "com.sec.android.app.popupcalculator:id/calc_keypad_btn_equal").click()

# 메시지 앱 실행
driver.activate_app("com.inzisoft.ods.meritz.itest")                            # 액티베이트 앱 (앱 패키지명) 으로 앱 실행
driver.implicitly_wait(3)
driver.find_element(AppiumBy.ID, "com.samsung.android.messaging:id/action_new_composer").click()                # 새 메시지 아이콘 클릭
driver.implicitly_wait(5)
driver.find_element(AppiumBy.ID, "com.samsung.android.messaging:id/single_chat_menu_layout").click()            # 1:1 대화 클릭

# 갤러리 앱 실행

driver.implicitly_wait(5)
driver.activate_app("com.sec.android.gallery3d")
driver.swipe(100,100,100,400)
driver.implicitly_wait(5)

driver.activate_app("com.android.vending")                                                  # 구글 플레이 스토어 앱 실행
driver.implicitly_wait(3)
driver.find_element (AppiumBy.XPATH, '//android.widget.TextView[@text="검색"]').click()     # 검색 탭 클릭
driver.implicitly_wait(3)
text = "dragon ball"    
driver.find_element (AppiumBy.XPATH, '//android.widget.TextView[@text="앱 및 게임 검색"]').click()              # 검색 입력 창 클릭
driver.find_element (AppiumBy.XPATH, "//android.widget.EditText").send_keys(text)                              # 드래곤 볼 입력 
driver.press_keycode(66)                   # 엔터키 입력
driver.implicitly_wait(10)

# 검색 결과에서 드래곤볼 레전즈 클릭
driver.find_element(AppiumBy.ACCESSIBILITY_ID, "DRAGON BALL LEGENDS\nBandai Namco Entertainment Inc.\n액션\n롤플레잉\n액션 전략\n별표 평점: 4.6\n5,000만회 이상 다운로드됨\n").click()
driver.implicitly_wait(10)
driver.find_element(AppiumBy.ACCESSIBILITY_ID, "설치").click()         # 설치 버튼 클릭

for i in range(2):                                  # 밑에서 위로 스와이프 동작 2회 실시
    driver.swipe (31,1104,13,118)
    driver.implicitly_wait(10)

# 앱 종료
driver.terminate_app("com.android.vending")
driver.terminate_app("com.samsung.android.messaging")

driver.activate_app("com.inzisoft.demo.itest")                  # 앱 실행
driver.implicitly_wait(10)

# 로그인 화면에서 id 입력 필드에 id 입력
driver.find_element (AppiumBy.XPATH, '(//android.widget.FrameLayout[@resource-id="com.inzisoft.demo.itest:id/fragment_container_view"])[2]/androidx.compose.ui.platform.ComposeView/android.view.View/android.view.View[1]/android.view.View/android.widget.EditText[4]').send_keys("apodE2R335가나")

# 로그인 화면에서 pw 입력 필드 클릭
driver.find_element (AppiumBy.XPATH, '(//android.widget.FrameLayout[@resource-id="com.inzisoft.demo.itest:id/fragment_container_view"])[2]/androidx.compose.ui.platform.ComposeView/android.view.View/android.view.View[1]/android.view.View/android.widget.EditText[5]').click()
#driver.find_element (AppiumBy.XPATH, '//android.widget.EditText').send_keys ("!@2334dlfdl일이")

# pw 보안 키패드에서 특정 글자 클릭 후 입력확인 버튼 클릭
driver.find_element (AppiumBy.ACCESSIBILITY_ID, '1 버튼').click()
driver.find_element (AppiumBy.ACCESSIBILITY_ID, 'q 버튼').click()
driver.find_element (AppiumBy.ACCESSIBILITY_ID, 'm 버튼').click()
driver.find_element (AppiumBy.ACCESSIBILITY_ID, '입력확인 버튼').click()
driver.implicitly_wait(10)
print(driver.page_source)                                               # 현재 페이지 소스 출력

# 인증 실패 팝업 발생시 해당 팝업 클릭 (해당 xpath 만 clickable = True 인데도, 동작 안됨)
#driver.find_element (AppiumBy.XPATH, '//androidx.compose.ui.platform.ComposeView/android.view.View/android.view.View').click()

# PointerInput 생성
pointer = PointerInput(kind="touch", name="touch")
actions = ActionBuilder(driver, mouse=pointer)

# clickable = True 인 xpath 동작되지 않아, 특정 좌표 클릭 (예: x=150, y=300)
actions.pointer_action.move_to_location(1000, 693)                          # 앱 내의 커스텀 알림 팝업의 [확인] 부분의 좌표 위치
actions.pointer_action.click()
actions.perform()                                   # 실제 포인터 이동 및 터치 동작 수행
driver.implicitly_wait(10)
# 로그인 버튼 클릭
driver.find_element (AppiumBy.XPATH, '//android.widget.TextView[@text="로그인"]').click()
driver.implicitly_wait(10)

# 앱 종료
driver.terminate_app("com.inzisoft.demo.itest")

driver.activate_app ("com.sec.android.app.camera")                                              # 카메라 앱 실행
driver.implicitly_wait(10)
driver.find_element (AppiumBy.ID, 'com.sec.android.app.camera:id/normal_center_button').click()         # 촬영 버튼 클릭
driver.implicitly_wait(10)

driver.find_element (AppiumBy.ACCESSIBILITY_ID, '후면 카메라로 전환').click()                           # 후면 카메라로 전환 버튼 클릭
driver.implicitly_wait(10)
driver.find_element (AppiumBy.ID, 'com.sec.android.app.camera:id/ripple_effect_view').click()               # 갤러리 버튼 클릭

driver.find_element (AppiumBy.ACCESSIBILITY_ID, '삭제').click()                                       # 삭제 (휴지통 아이콘) 버튼 클릭  
driver.implicitly_wait(10)
driver.find_element (AppiumBy.ID, 'android:id/button1').click()                                     # 팝업의 '휴지통으로 이동' 클릭
driver.implicitly_wait(10)
driver.find_element (AppiumBy.ACCESSIBILITY_ID, '상위 메뉴로 이동').click()                          # '<' 버튼 클릭   

driver.swipe (576,1070,638,281)                                                                     # 밑에서 위 방향으로 스와이프
driver.activate_app ("com.sec.android.app.camera")                                              # 카메라 앱 실행
driver.implicitly_wait(10)
driver.find_element (AppiumBy.ACCESSIBILITY_ID, '전면 카메라로 전환').click()                       # 전면 카메라로 전환 버튼 클릭
driver.implicitly_wait(10)
driver.terminate_app ("com.sec.android.app.camera")                                             # 카메라 앱 종료

driver.activate_app ("com.samsung.android.app.notes")
WebDriverWait(driver,2)
driver.find_element (AppiumBy.ACCESSIBILITY_ID, '노트 작성').click()
WebDriverWait(driver,2)

now = datetime.now()
format_datetime = now.strftime ("%Y-%m-%d %H:%M:%S")
format_datetime = str(format_datetime)

print (format_datetime)
WebDriverWait(driver,2)
driver.hide_keyboard()
driver.find_element (AppiumBy.XPATH, '//android.widget.RelativeLayout[@resource-id="com.samsung.android.app.notes:id/main_layout_container"]/android.widget.ScrollView/android.view.View').click()
WebDriverWait(driver,1)
driver.find_element (AppiumBy.XPATH, '//android.view.View[@content-desc=" "]').send_keys (format_datetime)
driver.press_keycode(66) 

driver.quit()
