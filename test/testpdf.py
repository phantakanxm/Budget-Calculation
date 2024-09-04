import tkinter as tk
from tkinter import filedialog

def attach_file():
    file_path = filedialog.askopenfilename()
    if file_path:
        # แสดงผลชื่อไฟล์ที่เลือกใน Label
        file_label.config(text=f"Selected file: {file_path}")

# สร้างหน้าต่างหลัก
root = tk.Tk()
root.title("Attach and Display File Example")

# สร้างปุ่ม Attach File
attach_button = tk.Button(root, text="Attach File", command=attach_file)
attach_button.pack(pady=20)

# Label สำหรับแสดงผลชื่อไฟล์
file_label = tk.Label(root, text="")
file_label.pack(pady=10)

# รันโปรแกรม
root.mainloop()
