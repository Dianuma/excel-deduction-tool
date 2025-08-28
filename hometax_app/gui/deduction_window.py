import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.filedialog import askopenfilename
from handler.excel_handler import excel_handler
from temp_data import temp_data as td
from handler.web_handler import web_handler
from config import *
import threading

class DeductionWindow(tk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.master = master

        '''상단 프레임'''
        # 프레임
        self.frm1 = tk.LabelFrame(self, text="공제 불공제 변경", pady=15, padx=15)   # pad 내부
        self.frm1.grid(row=0, column=0, pady=10, padx=10, sticky="nswe") # pad 내부

        '''요소 생성'''
        # 자료 선택
        self.lbl1 = tk.Label(self.frm1, text='자료가 들어있는 엑셀 파일을 선택해 주세요.')
        self.listbox1 = tk.Listbox(self.frm1, width=40, height=1)

        # 파일 선택 버튼
        self.select_button = tk.Button(self.frm1, text="찾아보기", width=12, command=self.search_file) 

        # 진행 상황 표시용 리스트박스
        self.lbl2 = tk.Label(self.frm1, text='진행상황.')
        self.listbox2 = tk.Listbox(self.frm1, width=40)

        # 진행률 프로그래스 바 
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.frm1, variable=self.progress_var, maximum=100, length=200, mode="determinate")
        self.progress_label = tk.Label(self.frm1, text="0%")

        '''버튼을 위한 서브 프레임'''
        self.button_frame = tk.Frame(self.frm1)
        self.button_frame.grid(row=3, column=3, rowspan=3, sticky="n", pady=5)

        # 버튼 폭 통일 (width=15 정도 추천)
        self.start_button = tk.Button(
            self.button_frame, text="변경 실행", width=12, 
            command=lambda: threading.Thread(target=web_handler.start_deduction_process, args=(self.update_progress,)).start()
            )
        self.site_button1 = tk.Button(self.button_frame, text="변경 사이트", width=12, command=lambda: web_handler.change_site_url(HOMETAX_URL_DEDUCTION_CHANGE))
        self.site_button2 = tk.Button(self.button_frame, text="조회 사이트", width=12, command=lambda: web_handler.change_site_url(HOMETAX_URL_DEDUCTION_CHECK))

        # 버튼 배치 (패킹)
        self.start_button.pack(pady=2)
        self.site_button1.pack(pady=2)
        self.site_button2.pack(pady=2)
        
        # 스크롤바 - 기능 연결
        self.scrollbar = tk.Scrollbar(self.frm1, orient="vertical")
        self.scrollbar.config(command=self.listbox2.yview)

        self.scrollbar_x = tk.Scrollbar(self.frm1, orient="horizontal")
        self.scrollbar_x.config(command=self.listbox2.xview)

        self.listbox2.config(yscrollcommand=self.scrollbar.set)
        self.listbox2.config(xscrollcommand=self.scrollbar_x.set)

        '''요소 배치'''
        # 상단 프레임
        self.lbl1.grid(row=0, column=1, columnspan=2, sticky="w")
        self.lbl2.grid(row=2, column=1, columnspan=2, sticky="w")
        self.listbox1.grid(row=1, column=1, columnspan=2, sticky="wens")
        self.listbox2.grid(row=3, column=1, rowspan=2, sticky="wens")
        self.scrollbar.grid(row=3, column=2, rowspan=2, sticky="ns")
        self.scrollbar_x.grid(row=5, column=1, sticky="ew")
        self.progress_bar.grid(row=6, column=1, columnspan=2, sticky="wens")
        self.progress_label.grid(row=7, column=1, columnspan=2, sticky="w")
        self.select_button.grid(row=1, column=3)


        '''하단 프레임'''
        # 프레임
        self.frm2 = tk.Frame(self, pady=10)
        self.frm2.grid(row=1, column=0, pady=10)

        # 요소 생성
        self.refresh_button = tk.Button(self.frm2, text="초기화", width=8, command=self.reset)
        self.next_button = tk.Button(self.frm2, text="이전", width=8, command=self.master.login_frame)
        self.quit_button = tk.Button(self.frm2, text="종료", width=8, command=self.master.all_quit)

        # 배치
        self.refresh_button.grid(row=0, column=3, sticky="wes")
        self.next_button.grid(row=0, column=6, sticky="ws")
        self.quit_button.grid(row=0, column=0, sticky="es")
        '''실행'''
        self.set_display()
    def set_display(self):
        self.listbox1.delete(0, "end")
        self.listbox2.delete(0, "end")
        if td.deduction_file is not None:
            self.listbox1.insert(0, td.deduction_file)
        for log in td.deduction_log_data:
            self.listbox2.insert(0, log[0])
            if log[1]:
                self.listbox2.itemconfig(0, {'fg':'red'})
    def search_file(self):  # 파일 검색
        try:
            if td.deduction_file_path:
                filename = askopenfilename(initialdir=td.get_deduction_file_path(), filetypes=(("Excel files", ".xlsx .xls"), ('All files', '*.*')))
            else:
                filename = askopenfilename(initialdir="./", filetypes=(("Excel files", ".xlsx .xls"), ('All files', '*.*')))
            self.listbox1.delete(0, "end")
            self.listbox1.insert(0, filename)
            excel_handler.load_data(filename)
        except Exception as e:
            messagebox.showerror("Error", f"파일 선택 중 오류 발생: {e}")
    def reset(self):  # 초기화
        try:
            reply = messagebox.askyesno("초기화", "정말로 초기화 하시겠습니까?")
            if reply:
                self.listbox1.delete(0, "end")
                self.listbox2.delete(0, "end")
                excel_handler.reset_deduction_data()
                td.reset_deduction_file()
                messagebox.showinfo("Success", "초기화 되었습니다.")
        except Exception as e:
            messagebox.showerror("Error", f"초기화 중 오류 발생: {e}")
    def update_progress(self, text, progress, isFail:bool=False):  # 진행률 업데이트
        td.deduction_log_data.append([text, isFail])
        self.listbox2.insert(0, text)
        if isFail:
            self.listbox2.itemconfig(0, {'fg':'red'})
        self.progress_var.set(progress)
        self.progress_label.config(text="")
        self.progress_label.config(text=f"{progress:.1f}%")
        self.update_idletasks()
