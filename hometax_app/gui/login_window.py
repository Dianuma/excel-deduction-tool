import tkinter as tk
from tkinter import messagebox
from handler.excel_handler import excel_handler
from temp_data import temp_data as td
from handler.web_handler import web_handler
from gui.ui_helper import select_file
from config import *

class LoginWindow(tk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.master = master
        self.excel_handler = excel_handler
        self.web_handler = web_handler

        '''상단 프레임 (로그인)'''
        # 프레임 생성 및 배치
        self.frm1 = tk.LabelFrame(self, text="로그인", pady=15, padx=15)
        self.frm1.grid(row=0, column=0, pady=10, padx=10, sticky="nswe")

        # 요소 생성
        self.lbl1 = tk.Label(self.frm1, text='회사의 ID, password가 있는 엑셀 파일을 선택해 주세요.')
        self.lbl2 = tk.Label(self.frm1, text='로그인 하고자 하는 회사를 선택해 주세요.')

        self.listbox1 = tk.Listbox(self.frm1, width=40, height=1)
        self.listbox2 = tk.Listbox(self.frm1, width=40)
        self.listbox2.bind("<Double-Button-1>", self.login)

        self.select_button = tk.Button(self.frm1, text="찾아보기", width=12, command=self.search_file)

        self.scrollbar = tk.Scrollbar(self.frm1)
        self.scrollbar.config(command=self.listbox2.yview)
        self.listbox2.config(yscrollcommand=self.scrollbar.set)

        # 배치
        self.lbl1.grid(row=0, column=1, columnspan=2, sticky="w")
        self.listbox1.grid(row=1, column=1, columnspan=2, sticky="wens")
        self.lbl2.grid(row=2, column=1, columnspan=2, sticky="w")
        self.listbox2.grid(row=3, column=1, rowspan=2, sticky="wens")
        self.scrollbar.grid(row=3, column=2, rowspan=2, sticky="wens")
        self.select_button.grid(row=1, column=3)

        '''하단 프레임 (종료, 초기화, 다음)'''
        # 프레임 생성 및 배치
        self.frm2 = tk.Frame(self, pady=10)
        self.frm2.grid(row=1, column=0, pady=10)
        
        # 요소 생성
        self.quit_button = tk.Button(self.frm2, text="종료", width=8, command=self.all_quit)
        self.refresh_button = tk.Button(self.frm2, text="초기화", width=8, command=self.reset)
        self.next_button = tk.Button(self.frm2, text="다음", width=8, command=self.master.deduction_frame)

        # 배치
        self.quit_button.grid(row=0, column=0, sticky="es")
        self.refresh_button.grid(row=0, column=3, sticky="wes")
        self.next_button.grid(row=0, column=6, sticky="ws")
        '''실행'''
        self.set_display()

    def set_display(self):
            self.listbox1.delete(0, "end")
            self.listbox2.delete(0, "end")
            if td.id_file is not None:
                self.listbox1.insert(0, td.id_file)
                self.display_company_list()
    def search_file(self, event = None):  # 파일 검색
        self.excel_handler.load_ids(select_file(td.get_id_file_path()))
        self.set_display()
    def display_company_list(self):  # 회사 목록 표시
        try:
            for id in self.excel_handler.id_data:
                self.listbox2.insert("end", self.get_display_text(id))
        except Exception as e:
            messagebox.showerror("Error", f"회사 목록 표시 중 오류 발생: {e}")
    def get_display_text(self, id):  # 회사 목록 표시 텍스트 표준화
        return f"{id} : {self.excel_handler.id_data[id][0]}"
    def login(self, event = None):  # 로그인
        try:
            selected_company = self.listbox2.get(self.listbox2.curselection())
            self.web_handler.hometax_login(selected_company.split(" : ")[0])
        except Exception as e:
            messagebox.showerror("Error", f"로그인 중 오류 발생: {e}")
    def reset(self, event = None):  # 초기화
        try:
            reply = messagebox.askyesno("초기화", "정말로 초기화 하시겠습니까?")
            if reply:
                self.listbox1.delete(0, "end")
                self.listbox2.delete(0, "end")
                self.excel_handler.reset_id_data()
                self.excel_handler.reset_deduction_data()
                td.reset_id_file()
                td.reset_deduction_file()
                messagebox.showinfo("Success", "초기화 되었습니다.")
        except:
            messagebox.showinfo("Error", "오류가 발생했습니다.")
    def all_quit(self, event = None):
        self.master.all_quit()