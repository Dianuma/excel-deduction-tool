import os

class TempData:
    def __init__(self):
        self.id_file = None
        self.deduction_file = None
        self.id_file_path = None
        self.deduction_file_path = None
        self.total_deduction = 0
        self.changed_deduction = 0
        self.curr_page = 0
        self.last_processed_row = 0
        self.deduction_log_data = []
        self.error_idx = []

    def reset_deduction_file(self):
        self.deduction_file = None
        self.total_deduction = 0
        self.changed_deduction = 0
        self.curr_page = 0
        self.last_processed_row = 0
        self.deduction_log_data.clear()
        self.error_idx.clear()
    def reset_id_file(self):
        self.id_file = None
    def set_id_file(self, filename):
        self.id_file = filename
        self.id_file_path = self._get_file_path(filename)
    def set_deduction_file(self, filename):
        self.deduction_file = filename
        self.deduction_file_path = self._get_file_path(filename)
    def get_id_file_path(self):
        return self.id_file_path
    def get_deduction_file_path(self):
        return self.deduction_file_path
    def _get_file_path(self, file_path):
        if file_path is None:
            return os.getcwd()
        dir_path = os.path.dirname(file_path)
        return dir_path if os.path.isdir(dir_path) else os.getcwd()
temp_data = TempData()