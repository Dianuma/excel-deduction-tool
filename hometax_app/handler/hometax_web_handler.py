from config import *
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.alert import Alert
from time import sleep
from handler.excel_handler import excel_handler
from temp_data import temp_data as td

class HometaxWebHandler:
    def __init__(self):
        self.driver = self.open_chrome()

    # ───────────── 브라우저 관련 ─────────────
    def open_chrome(self):  # 크롬 드라이버 오픈
        options = Options()

        for arg in CHROME_OPTIONS_ARGS:
            options.add_argument(arg)

        options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])  # 자동화 표시 제거
        options.add_experimental_option("useAutomationExtension", False)  # 자동화 확장 비활성화
        options.add_experimental_option("detach", True)

        driver = webdriver.Chrome(options=options)
        driver.get(url=HOMETAX_URL_MAIN)
        return driver
    def close_chrome(self):  # 크롬 드라이버 종료
        if self.driver:
            self.driver.quit()
    def change_site_url(self, url):  # 홈택스 사이트 URL 변경
        if self.driver.current_url != url:
            self.driver.get(url)

    # ───────────── 홈택스 관련 ─────────────
    def login_hometax(self, selected_company):  # 홈택스 로그인
        id_data = excel_handler.id_data.get(int(selected_company), None)

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
        # self.driver.switch_to.default_content()
    def deduction_change_process(self, func):  # 공제/불공제 변경
        total_rows = len(excel_handler.deduction_data)
        self.click_by_script(XPATH_ALL_SELECT_CHECKBOX)
        for i, data in enumerate(excel_handler.deduction_data):
            self.process_row(i, data, func)
            if self.is_last_row(i, total_rows):
                func(text=f"변경이 완료되었습니다.", progress=100)
                for error_idx in td.error_idx:
                    error_data = excel_handler.deduction_data[error_idx]
                    func(text=f"{error_data[0]} : {error_data[2]}의 엑셀 상 데이터 {error_data[4]}", progress=100, isFail=True)

                error_count = len(td.error_idx)
                func(text=f"실패한 항목 수 - {error_count:,}", progress=100, isFail=True)
                func(text=f"총 공제금액 - {excel_handler.total_deduction:,}", progress=100)
                self.save_changes()
            elif self.is_page_end(i):
                self.save_changes()
                self.change_page()
                self.click_by_script(XPATH_ALL_SELECT_CHECKBOX)

    # ───────────── 헬퍼 함수 ─────────────
    def get_cell(self, row, col):  # 요소 XPATH로 텍스트 가져오기
        return self.driver.find_element(By.XPATH, XPATH_DEDUCTION_CHANGE[col].format(row=row))
    def match_row_with_data(self, row, data):  # 행의 데이터와 엑셀 데이터를 비교
        day = self.get_cell(row, "day").text.replace(".", "-").strip()
        franchise_id = self.get_cell(row, "franchise_id").text.strip()
        name = self.get_cell(row, "name").text.strip()
        total = self.get_cell(row, "total").text.replace(",", "").strip()
        return day == str(data[0]) and franchise_id == str(data[1]) and name == str(data[2]) and total == str(int(data[3]))
    def click_by_script(self, xpath: str):  # 스크립트를 통한 클릭
        while True:
            try:
                selected = self.driver.find_element(By.XPATH, xpath)
                self.driver.execute_script(CLICK_SCRIPT, selected)
                break
            except:
                sleep(0.1)
        return True
    def get_pagination_xpath(self, kind: str = None, page_offset: int = None) -> str:  # 페이지 네비게이션 XPATH 가져오기
            """
            kind = "first", "prev", "next", "last" 중 하나
            page_offset = 현재 구간 내에서의 페이지 위치 (0~9)
            """
            if kind:
                idx  = XPATH_PAGE_NAVIGATION["index"].get(kind)
            elif page_offset is not None:
                idx = XPATH_PAGE_NAVIGATION["index"]["page_start"] + page_offset
            else:
                return None
            return XPATH_PAGE_NAVIGATION["base"].format(idx)
    def change_page(self):  # 페이지 이동
        td.curr_page += 1
        if td.curr_page % PAGE_BLOCK_SIZE == 0:  # 구간 이동
            self.click_by_script(self.get_pagination_xpath(kind="next"))
        else:  # 구간 내 페이지 이동
            self.click_by_script(self.get_pagination_xpath(page_offset=td.curr_page % PAGE_BLOCK_SIZE))
        sleep(0.5)
        self.alert_check()
    def change_deduction(self, xpath, value, data, i, func):  # 공제/불공제 변경
        try:
            Select(self.driver.find_element(By.XPATH, xpath)).select_by_index(value)
            if value == 0:
                td.total_deduction += int(data[3])
            func(text=f"{i+1} - {data[2]} : {data[4]}", progress=self.get_progress_percentage(i))
        except Exception as e:
            td.error_idx.append(i)
            func(text=f"{i+1} - {data[0]} {data[2]}의 공제 불공제 변경에 실패했습니다. - {data[4]}", progress=self.get_progress_percentage(i), isFail=True)
    def alert_check(self):  # 알림창 확인
        try:
            alert = Alert(self.driver)
            alert.accept()
        except:
            pass
    def process_row(self, i, data, func):  # 행 처리
        table_order = i % ROW_PER_PAGE + 1
        if self.match_row_with_data(table_order, data):
            xpath = XPATH_DEDUCTION_CHANGE["select"].format(row=table_order)
            value = 0 if data[4] == "공제" else 1
            self.change_deduction(xpath, value, data, i, func)
        else:
            td.error_idx.append(i)
            func(text=f"{i+1} - {data[0]} {data[2]}의 값이 홈택스와 다릅니다. - {data[4]}", progress=self.get_progress_percentage(i), isFail=True)
    def get_progress_percentage(self, i):  # 진행률 계산
        return (i + 1) / len(excel_handler.deduction_data) * 100
    def save_changes(self):  # 변경 사항 저장
        self.click_by_script(XPATH_DEDUCTION_CHANGE["change_button"])
        sleep(0.5)
        self.alert_check()
    def is_page_end(self, i):  # 페이지 끝 확인
        return (i + 1) % ROW_PER_PAGE == 0
    def is_last_row(self, i, total_rows):  # 마지막 행 확인
        return i == total_rows - 1

hometax_web_handler = HometaxWebHandler()