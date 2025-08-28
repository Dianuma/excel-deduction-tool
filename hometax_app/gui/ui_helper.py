from tkinter.filedialog import askopenfilename
from tkinter import messagebox

def select_file(initial_dir):
    try:
        return askopenfilename(initialdir=initial_dir, filetypes=(("Excel files", ".xlsx .xls"), ('All files', '*.*')))
    except Exception as e:
        messagebox.showerror("Error", f"파일 선택 중 오류 발생: {e}")
