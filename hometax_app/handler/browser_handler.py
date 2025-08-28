from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from config import *

class BrowserHandler:
    def __init__(self):
        self.driver = None

    def open_chrome(self):  # 크롬 드라이버 오픈
        options = Options()
        for arg in CHROME_OPTIONS_ARGS:
            options.add_argument(arg)
        options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
        options.add_experimental_option("useAutomationExtension", False)
        options.add_experimental_option("detach", True)

        self.driver = webdriver.Chrome(options=options)
        self.driver.get(url=HOMETAX_URL_MAIN)
        return self.driver
    def change_site_url(self, url):  # 홈택스 사이트 URL 변경
        if self.driver.current_url != url:
            self.driver.get(url)
    def close_chrome(self):  # 크롬 드라이버 종료
        if self.driver:
            self.driver.quit()