from config import *
from selenium.webdriver.common.by import By
from time import sleep
from handler.excel_handler import excel_handler
from handler.selenium_helper import click_by_script, alert_check
from selenium.webdriver.support.select import Select
from temp_data import temp_data as td

class HomeTaxDeductionHandler:
    def __init__(self, driver):
        self.driver = driver
        self.excel_handler = excel_handler
        self.func = None

    def deduction_change_process(self, progress_callback):  # 공제/불공제 변경
        self.func = progress_callback
        total_rows = len(excel_handler.deduction_data)
        click_by_script(self.driver, XPATH_ALL_SELECT_CHECKBOX)
        for i, data in enumerate(excel_handler.deduction_data):
            self.process_row(i, data)
            if self.is_last_row(i, total_rows):
                self.func(text=f"변경이 완료되었습니다.", progress=100)
                for error_idx in td.error_idx:
                    error_data = excel_handler.deduction_data[error_idx]
                    self.func(text=f"{error_data[0]} : {error_data[2]}의 엑셀 상 데이터 {error_data[4]}", progress=100, isFail=True)

                error_count = len(td.error_idx)
                self.func(text=f"실패한 항목 수 - {error_count:,}", progress=100, isFail=True)
                self.func(text=f"총 공제금액 - {excel_handler.total_deduction:,}", progress=100)
                self.save_changes()
            elif self.is_page_end(i):
                self.save_changes()
                self.change_page()
                click_by_script(self.driver, XPATH_ALL_SELECT_CHECKBOX)

    def get_cell(self, row, col):  # 요소 XPATH로 텍스트 가져오기
        return self.driver.find_element(By.XPATH, XPATH_DEDUCTION_CHANGE[col].format(row=row))
    def match_row_with_data(self, row, data):  # 행의 데이터와 엑셀 데이터를 비교
        day = self.get_cell(row, "day").text.replace(".", "-").strip()
        franchise_id = self.get_cell(row, "franchise_id").text.strip()
        name = self.get_cell(row, "name").text.strip()
        total = self.get_cell(row, "total").text.replace(",", "").strip()
        return day == str(data[0]) and franchise_id == str(data[1]) and name == str(data[2]) and total == str(int(data[3]))
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
            click_by_script(self.driver, self.get_pagination_xpath(kind="next"))
        else:  # 구간 내 페이지 이동
            click_by_script(self.driver, self.get_pagination_xpath(page_offset=td.curr_page % PAGE_BLOCK_SIZE))
        sleep(0.5)
        alert_check(self.driver)
    def change_deduction(self, xpath, value, data, i):  # 공제/불공제 변경
        try:
            Select(self.driver.find_element(By.XPATH, xpath)).select_by_index(value)
            if value == 0:
                td.total_deduction += int(data[3])
            self.func(text=f"{i+1} - {data[2]} : {data[4]}", progress=self.get_progress_percentage(i))
        except Exception as e:
            td.error_idx.append(i)
            self.func(text=f"{i+1} - {data[0]} {data[2]}의 공제 불공제 변경에 실패했습니다. - {data[4]}", progress=self.get_progress_percentage(i), isFail=True)
    def process_row(self, i, data):  # 행 처리
        table_order = i % ROW_PER_PAGE + 1
        if self.match_row_with_data(table_order, data):
            xpath = XPATH_DEDUCTION_CHANGE["select"].format(row=table_order)
            value = 0 if data[4] == "공제" else 1
            self.change_deduction(xpath, value, data, i)
        else:
            td.error_idx.append(i)
            self.func(text=f"{i+1} - {data[0]} {data[2]}의 값이 홈택스와 다릅니다. - {data[4]}", progress=self.get_progress_percentage(i), isFail=True)
    def get_progress_percentage(self, i):  # 진행률 계산
        return (i + 1) / len(excel_handler.deduction_data) * 100
    def save_changes(self):  # 변경 사항 저장
        click_by_script(self.driver, XPATH_DEDUCTION_CHANGE["change_button"])
        sleep(0.5)
        alert_check(self.driver)
    def is_page_end(self, i):  # 페이지 끝 확인
        return (i + 1) % ROW_PER_PAGE == 0
    def is_last_row(self, i, total_rows):  # 마지막 행 확인
        return i == total_rows - 1
