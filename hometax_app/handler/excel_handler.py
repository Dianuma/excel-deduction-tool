from temp_data import temp_data as td
import pandas as pd
import math

class ExcelHandler:
    def __init__(self):
        self.deduction_data = []
        self.id_data = {}
        self.total_deduction = 0

    def open_excel(self, filename):  # 엑셀 파일 열기
        try:
            if filename.endswith('.xls'):
                df = pd.read_excel(filename, engine='xlrd')
            elif filename.endswith('.xlsx'):
                df = pd.read_excel(filename, engine='openpyxl')
            else:
                raise ValueError("지원하지 않는 파일 형식입니다. (.xls 또는 .xlsx)")
            return df
        except Exception as e:
            print(f"Excel 로드 실패: {e}")
            return None
    def load_ids(self, filename):  # ID 데이터 로드
        td.id_file_path = filename
        df = self.open_excel(filename)
        if df is None:
            return []

        all_values = {}

        for _, row in df.iterrows():
            row_value = [self.process_data(cell) for cell in row]
            if row_value[0] != "None" and row_value[4] is not None:
                all_values[row_value[0]] = [row_value[1], row_value[2], row_value[3], row_value[4]]

        self.id_data = all_values
    def load_data(self, filename):  # 데이터 로드
        td.deduction_file_path = filename
        df = self.open_excel(filename)
        if df is None:
            return []

        all_values = []
        for _, row in df.iterrows():
            row_value = [str(cell) for cell in row]
            row_input = [self.process_data(row_value[i]) for i in [0, 3, 4, 8, 12]] # 승인일자, 가맹점 사업자번호, 가맹점명, 합계금액, 공제여부
            all_values.append(row_input)

        self.deduction_data = all_values[1:]  # 헤더 제외
        self.set_total_deduction()
    def process_data(self, cell) :  # 데이터 전처리
        if isinstance(cell, float):
            if math.isnan(cell):  # 비어 있는 셀 처리
                return None       # 또는 원하는 기본값
            return int(cell)
        elif isinstance(cell, str):
            return cell.strip()
        return cell
    def reset_id_data(self):  # ID 데이터 초기화
        self.id_data = {}
    def reset_deduction_data(self):  # 공제 데이터 초기화
        self.deduction_data = []
        self.total_deduction = 0
    def set_total_deduction(self):  # 총 공제액 설정
        for data in self.deduction_data:
            if data[4] == "공제":
                self.total_deduction += int(data[3])
excel_handler = ExcelHandler()