import os

class TempData:
    def __init__(self):
        self.id_file_path = None
        self.deduction_file_path = None
        self.total_deduction = 0
        self.changed_deduction = 0
        self.curr_page = 0
        self.deduction_log_data = []
        self.error_idx = []

    def reset(self):
        self.deduction_file_path = None
        self.total_deduction = 0
        self.changed_deduction = 0
        self.curr_page = 0
        self.last_processed_row = 0
        self.deduction_log_data = []
        self.error_idx = []
    def get_id_file_path(self):
        if os.path.isfile(self.id_file_path):
            return os.path.dirname(self.id_file_path)
        else:
            raise FileNotFoundError(f"ID 파일을 찾을 수 없습니다: {self.id_file_path}")
    def get_deduction_file_path(self):
        if os.path.isfile(self.deduction_file_path):
            return os.path.dirname(self.deduction_file_path)
        else:
            raise FileNotFoundError(f"공제 파일을 찾을 수 없습니다: {self.deduction_file_path}")

temp_data = TempData()