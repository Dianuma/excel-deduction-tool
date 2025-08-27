import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.filedialog import askopenfilename
from gui.login_window import LoginWindow
from handler.hometax_web_handler import hometax_web_handler as wh
from gui.deduction_window import DeductionWindow
class display_interface(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.title('공제 불공제 변경 프로그램')
        self.minsize(400, 300)  # 최소 사이즈
        self.resizable(False, False)
        self._frame = None
        self.ID_password = None
        self.selected_filename = None
        self.login_frame()

    def login_frame(self):
        self.switch_frame(LoginWindow)
    def deduction_frame(self):
        self.switch_frame(DeductionWindow)
    def switch_frame(self, frame_class):
        try:
            new_frame = frame_class(self)
            if hasattr(self, "_current_frame") and self._current_frame is not None:
                self._current_frame.grid_forget()  # 현재 프레임을 숨김
            self._current_frame = new_frame
            self._current_frame.grid(row=0, column=0, sticky="nsew")
        except Exception as e:
            messagebox.showinfo("Error", f"오류가 발생했습니다: {e}")
    def all_quit(self):
        try:
            reply = messagebox.askyesno("종료", "정말로 종료 하시겠습니까?")
            if reply:
                wh.close_chrome()
                self.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"종료 중 오류 발생: {e}")