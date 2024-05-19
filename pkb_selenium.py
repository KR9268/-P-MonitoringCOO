from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from win32 import win32gui
import time
import pyautogui
import pyperclip


def set_browser(url, headless=None, size_mini=None):
        #url = "https://www.net"
        if headless == None:
            driver = webdriver.Chrome()
            driver.get(url)
        else:
            opt = Options()
            opt.add_argument("--headless")
            driver = webdriver.Chrome(options=opt)
            driver.get(url)
        
        if size_mini == None:
            pass
        elif size_mini == 'Max':
            driver.maximize_window()
        else:
            driver.set_window_size(*size_mini)

        return driver


def box_input(driver, memo, type, address, input_text):
    '''
    By. (XPATH / CSS_SELECTOR / CLASS_NAME)
    example)box_input(driver, '본문내용추가', By.XPATH, '//*[@id="INPUT_72319310708"]')
    '''
    driver.find_element(type, address).send_keys(input_text)
    return memo

def box_clear(driver, memo, type, address):
    '''
    By. (XPATH / CSS_SELECTOR / CLASS_NAME)
    example)box_clear(driver, '내용삭제', By.XPATH, '//*[@id="INPUT_72319310708"]')
    '''
    driver.find_element(type, address).clear()
    return memo

def button_click(driver, memo, type, address):
    '''
    By. (XPATH / CSS_SELECTOR / CLASS_NAME)
    example)button_click(driver, '메일 전체선택', By.XPATH, '//*[@id="INPUT_72319310708"]')
    '''
    driver.find_element(type, address).click()
    return memo

def arguments_click(driver, memo, type, address):
    '''
    By. (XPATH / CSS_SELECTOR / CLASS_NAME)
    example)arguments_click(driver, '메일 전체선택', By.XPATH, '//*[@id="INPUT_72319310708"]')
    '''
    driver.execute_script("arguments[0].click();", driver.find_element(type, address))
    return memo

def get_text(driver, memo, type, address):
    '''
    By. (XPATH / CSS_SELECTOR / CLASS_NAME)
    example)get_text(driver, '텍스트얻기', By.XPATH, '//*[@id="INPUT_72319310708"]')
    '''
    return driver.find_element(type, address).text

def get_object(driver, memo, type, address):
    '''
    By. (XPATH / CSS_SELECTOR / CLASS_NAME)
    example)get_text(driver, '텍스트얻기', By.XPATH, '//*[@id="INPUT_72319310708"]')
    '''
    return driver.find_element(type, address)

def get_multi_elements(driver, memo, type, address):
    '''
    By. (XPATH / CSS_SELECTOR / CLASS_NAME)
    example)get_multi_elements(driver, '조건에맞는 요소 얻기', By.XPATH, '//*[@id="INPUT_72319310708"]')
    '''
    return driver.find_elements(type, address)

def get_foreground_window_text(): #현재 윈도우 확인용
    '''
    win32gui import 필요
    '''
    title = ""
    current_hwnd = win32gui.GetForegroundWindow()
    if current_hwnd is not None:
        title = win32gui.GetWindowText(win32gui.GetForegroundWindow())
    return title

def knox_mail_send_html(driver2, mail_subject, mail_body, mail_address=None, attach=None):
    '''
        knox_mail_send_html(driver2, mail_subject, mail_body, mail_address=None, attach=None)
        
        from selenium import webdriver
        from selenium.webdriver.common.by import By
        from selenium.webdriver.chrome.options import Options
        from selenium.webdriver.common.keys import Keys
        import chromedriver_autoinstaller
        import time
        import pyautogui
        import pyperclip
        from win32 import win32gui
    '''
    # 프레임전환하여 메일쓰기 버튼 클릭
    driver2.implicitly_wait(10)
    driver2.switch_to.frame(driver2.find_element(By.CSS_SELECTOR, "#pt-mail-iframe"))
    driver2.implicitly_wait(10)
    button_click(driver2, '메일쓰기클릭', By.CSS_SELECTOR,"body > div.content-container > div.lnb-area.ui-resizable > div.content > div.lnb-head > div > div > a > span > i")

    # 메일쓰기 창으로 이동
    time.sleep(2)
    driver2.switch_to.window(driver2.window_handles[1])
    driver2.implicitly_wait(10)
    box_input(driver2, '메일제목', By.CSS_SELECTOR, "#input-title",mail_subject)

    #발송할 주소가 없으면  나에게 쓰고 있으면 받아서 쓴다
    driver2.implicitly_wait(10)
    if mail_address is None:
        # 수신처지정(나에게 쓰기)
        driver2.implicitly_wait(10)
        button_click(driver2, '나에게쓰기 클릭', By.CSS_SELECTOR,"#recipients-area > div > div.form-add-receiver.folding-content > div > div.pa-wrapper > div > div > button.button.default.add-myself > span")
        alert_check = driver2.find_elements(By.CSS_SELECTOR, "#alertDialog")
        if len(alert_check) > 0:
            button_click(driver2, '팝업닫기', By.CSS_SELECTOR,"#alertDialog > div.pop-btn > button")
    else:
        # 수신처지정(수신처지정)
        if type(mail_address) == str:
            driver2.implicitly_wait(10)
            box_input(driver2, '메일주소 입력', By.XPATH, '/html/body/div/div[2]/div[2]/div/form/table/tbody/tr[2]/td[2]/div/div[2]/div/div[1]/fieldset/div/input', mail_address)
            driver2.implicitly_wait(10)
            box_input(driver2, '엔터', By.XPATH, '/html/body/div/div[2]/div[2]/div/form/table/tbody/tr[2]/td[2]/div/div[2]/div/div[1]/fieldset/div/input', Keys.ENTER)   
            driver2.implicitly_wait(10)
            box_input(driver2, 'ESC', By.XPATH, '/html/body/div/div[2]/div[2]/div/form/table/tbody/tr[2]/td[2]/div/div[2]/div/div[1]/fieldset/div/input', Keys.ESCAPE)
        elif type(mail_address) == list:
            for each_mail_address in mail_address:
                driver2.implicitly_wait(10)
                box_input(driver2, '메일주소 입력', By.XPATH, '/html/body/div/div[2]/div[2]/div/form/table/tbody/tr[2]/td[2]/div/div[2]/div/div[1]/fieldset/div/input', each_mail_address)
                driver2.implicitly_wait(10)
                box_input(driver2, '엔터', By.XPATH, '/html/body/div/div[2]/div[2]/div/form/table/tbody/tr[2]/td[2]/div/div[2]/div/div[1]/fieldset/div/input', Keys.ENTER) 
                driver2.implicitly_wait(10)  
                box_input(driver2, 'ESC', By.XPATH, '/html/body/div/div[2]/div[2]/div/form/table/tbody/tr[2]/td[2]/div/div[2]/div/div[1]/fieldset/div/input', Keys.ESCAPE)

    # 파일첨부
    driver2.implicitly_wait(10)
    if attach is None:
        pass
    else:
        if type(attach) == str:
            button_click(driver2, '첨부버튼2', By.XPATH,"/html/body/div[1]/div[2]/div[2]/div/form/table/tbody/tr[3]/td/div/form/div/div[1]/span[2]/div")
            # 팝업창이 열릴때까지 기다린다
            window = None
            while window is None:  
                for _ in pyautogui.getAllWindows():
                    if _.title == '열기': 
                        window = _  
            while True:
                window.activate()
                if get_foreground_window_text()=='열기': 
                    break
            #첨부시작
            time.sleep(1)
            pyperclip.copy(attach)
            pyautogui.hotkey("ctrl", "v")
            pyautogui.hotkey("enter")
        elif type(attach) == list:
            for file_address in attach:
                button_click(driver2, '첨부버튼2', By.XPATH,"/html/body/div[1]/div[2]/div[2]/div/form/table/tbody/tr[3]/td/div/form/div/div[1]/span[2]/div")
                # 팝업창이 열릴때까지 기다린다
                window = None
                while window is None:  
                    for _ in pyautogui.getAllWindows():
                        if _.title == '열기': 
                            window = _  
                while True:
                    window.activate()
                    if get_foreground_window_text()=='열기': 
                        break
                   #첨부시작  
                time.sleep(1)
                pyperclip.copy(file_address)
                pyautogui.hotkey("ctrl", "v")
                pyautogui.hotkey("enter")
                time.sleep(1)
        else:
            pass

    button_click(driver2, 'html전환', By.XPATH, '/html/body/div/div[2]/div[2]/div/form/table/tbody/tr[5]/td/div[2]/div[2]/div[2]/ul/li[2]')
    box_clear(driver2, '기존내용지우기', By.XPATH, '/html/body/div/div[2]/div[2]/div/form/table/tbody/tr[5]/td/div[2]/div[2]/div[2]/div[2]/textarea')
    box_input(driver2, '메일본문 입력', By.XPATH, '/html/body/div/div[2]/div[2]/div/form/table/tbody/tr[5]/td/div[2]/div[2]/div[2]/div[2]/textarea', mail_body)

    driver2.switch_to.default_content()

    button_click(driver2, '발신버튼', By.CSS_SELECTOR,"body > div > div.win-header > div > div > button.button.default.blue.ng-scope")
    button_click(driver2, '발신 Yes', By.CSS_SELECTOR,"#sendExtMailRestrainedUserConfirmDialog > div.pop-btn > button.button.default.b > span")

    # 메일쓰기 프레임으로 돌아가기
    driver2.switch_to.window(driver2.window_handles[0])
