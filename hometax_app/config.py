# 홈택스 사이트 URL
HOMETAX_URL_MAIN = "https://www.hometax.go.kr/websquare/websquare.html?w2xPath=/ui/pp/index_pp.xml"
HOMETAX_URL_DEDUCTION_CHANGE = "https://hometax.go.kr/websquare/websquare.html?w2xPath=/ui/pp/index_pp.xml&tmIdx=46&tm2lIdx=4608020000&tm3lIdx=4608020100"
HOMETAX_URL_DEDUCTION_CHECK = "https://hometax.go.kr/websquare/websquare.html?w2xPath=/ui/pp/index_pp.xml&tmIdx=46&tm2lIdx=4608020000&tm3lIdx=4608020200"

# 홈택스 아이디 입력 XPATH
XPATH_LOGIN = {
    "id": '/html/body/div[1]/div[2]/div/div[1]/div/div[1]/div[2]/div[2]/div[3]/div/div[1]/div[1]/ul/li[1]/div/input',
    "pw": '/html/body/div[1]/div[2]/div/div[1]/div/div[1]/div[2]/div[2]/div[3]/div/div[1]/div[1]/ul/li[2]/div/input',
    "button": '/html/body/div[1]/div[2]/div/div[1]/div/div[1]/div[2]/div[2]/div[3]/div/div[1]/div[2]/a'
}

# 홈택스 주민번호 입력 XPATH
XPATH_RES_NO = {
    "front": '/html/body/div[6]/div[2]/div[1]/div/div/div[1]/div[2]/div/div[2]/div/ul/li[2]/div/input[1]',
    "back": '/html/body/div[6]/div[2]/div[1]/div/div/div[1]/div[2]/div/div[2]/div/ul/li[2]/div/input[2]',
    "button": '/html/body/div[6]/div[2]/div[1]/div/div/div[1]/div[2]/div/div[2]/div/div/input'
}

# 홈택스 공제 불공제 변경 XPATH
XPATH_ALL_SELECT_CHECKBOX = '/html/body/div[1]/div[2]/div/div/div[2]/div[4]/div[2]/div/div[1]/div/table/thead/tr/th[1]/input'

XPATH_DEDUCTION_CHANGE = {
    "day": "/html/body/div[1]/div[2]/div/div/div[2]/div[4]/div[2]/div/div[1]/div/table/tbody/tr[{row}]/td[2]",
    "franchise_id": "/html/body/div[1]/div[2]/div/div/div[2]/div[4]/div[2]/div/div[1]/div/table/tbody/tr[{row}]/td[5]",
    "name": "/html/body/div[1]/div[2]/div/div/div[2]/div[4]/div[2]/div/div[1]/div/table/tbody/tr[{row}]/td[6]",
    "total": "/html/body/div[1]/div[2]/div/div/div[2]/div[4]/div[2]/div/div[1]/div/table/tbody/tr[{row}]/td[10]",
    "select" : "/html/body/div[1]/div[2]/div/div/div[2]/div[4]/div[2]/div/div[1]/div/table/tbody/tr[{row}]/td[14]/div/div/select",
    "change_button" : "/html/body/div[1]/div[2]/div/div/div[2]/div[5]/div/span/input"
}

XPATH_PAGE_NAVIGATION = {
    "base" : "/html/body/div[1]/div[2]/div/div/div[2]/div[4]/div[3]/div/div/div[1]/ul/li[{}]/a",
    "index" : {
        "first" : 1,
        "prev" : 2,
        "page_start" : 3, # 3 ~ 12
        "next" : 13,
        "last" : 14
    }
}

CLICK_SCRIPT = "arguments[0].click();"

# 크롬 드라이버 옵션
CHROME_OPTIONS_ARGS = [
    "--disable-logging",
    "--log-level=3",
    "--no-sandbox",
    "--disable-gpu",
    "--disable-extensions",
    "--remote-debugging-port=0"
]

# 엑셀 관련 설정
EXCEL_FILE_TYPES = [("Excel files", ".xlsx .xls"), ("All files", "*.*")]

# 메시지
MSG_ERROR = "오류가 발생했습니다."
MSG_SUCCESS_INIT = "초기화 되었습니다."
MSG_CONFIRM_EXIT = "정말로 종료 하시겠습니까?"
MSG_CONFIRM_REFRESH = "정말로 초기화 하시겠습니까?"

# 기타 설정
PAGE_BLOCK_SIZE = 10  # 페이지 블록 당 페이지 수
ROW_PER_PAGE = 20  # 페이지 당 행 수