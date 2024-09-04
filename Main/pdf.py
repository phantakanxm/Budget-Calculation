import tkinter as tk
from tkinter import filedialog

def attach_file(file_label,btn):
    file_path = filedialog.askopenfilename()
    if file_path:
        # ทำสิ่งที่คุณต้องการกับไฟล์ที่เลือก
       file_label.config(text=f"แนบไฟล์แล้ว")
       btn.destroy()
    #    file_label.config(text=f"Selected file: {file_path}")



