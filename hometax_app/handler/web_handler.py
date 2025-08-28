from handler.browser_handler import BrowserHandler
from handler.hometax_login_handler import HomeTaxLoginHandler as HometaxLogin
from handler.hometax_deduction_handler import HomeTaxDeductionHandler as HometaxDeduction

class WebHandler:
    def __init__(self):
        self.browser = BrowserHandler()
        self.driver = self.browser.open_chrome()
        self.login_handler = HometaxLogin(self.driver)
        self.deduction_handler = HometaxDeduction(self.driver)

    def change_site_url(self, url):
        self.browser.change_site_url(url)
    def start_deduction_process(self, progress_callback):
        self.deduction_handler.deduction_change_process(progress_callback)
    def hometax_login(self, selected_company):
        self.login_handler.login_hometax(selected_company)
    def close(self):
        self.browser.close_chrome()

web_handler = WebHandler()