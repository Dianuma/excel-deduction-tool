from time import sleep
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.by import By
from config import *

def click_by_script(driver, xpath: str):  # 스크립트를 통한 클릭
    while True:
        try:
            selected = driver.find_element(By.XPATH, xpath)
            driver.execute_script(CLICK_SCRIPT, selected)
            break
        except:
            sleep(0.1)
    return True
def alert_check(driver):  # 알림창 확인
    try:
        alert = Alert(driver)
        alert.accept()
    except:
        pass