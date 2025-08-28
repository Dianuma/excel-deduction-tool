from config import *
from selenium.webdriver.common.by import By
from time import sleep
from handler.excel_handler import excel_handler

class HomeTaxLoginHandler:
    def __init__(self, driver):
        self.driver = driver
        self.excel_handler = excel_handler

    def login_hometax(self, selected_company):  # 홈택스 로그인
        id_data = self.excel_handler.id_data.get(int(selected_company), None)

        if not id_data:
            raise ValueError("선택한 회사의 ID 정보가 없습니다.")

        id, pw = id_data[1], id_data[2]
        res_no = str(id_data[3])
        res_no_front, res_no_back_first = res_no[:6], res_no[6]

        # ID 입력
        self.driver.find_element(By.XPATH, XPATH_LOGIN["id"]).send_keys(id)
        self.driver.find_element(By.XPATH, XPATH_LOGIN["pw"]).send_keys(pw)
        self.driver.find_element(By.XPATH, XPATH_LOGIN["button"]).click()

        # 주민번호 입력
        while True:
            try:
                self.driver.find_element(By.XPATH, XPATH_RES_NO["front"]).send_keys(res_no_front)
                self.driver.find_element(By.XPATH, XPATH_RES_NO["back"]).send_keys(res_no_back_first)
                self.driver.find_element(By.XPATH, XPATH_RES_NO["button"]).click()
                break
            except:
                sleep(0.1)
        return True