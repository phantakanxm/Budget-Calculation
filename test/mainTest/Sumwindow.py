import matplotlib.pyplot as plt
import numpy as np
import seaborn as sns

from tkinter import Tk, Canvas, Entry, Button
from tkinter import messagebox
from tkinter import *  
from tkinter import ttk
import tkcalendar as tkc
from datetime import date


from InputData import InputDataToSubject, InputDataToSum, InputDataToIncome #function ส่งข้อมูลไปยัง excel

from openpyxl import load_workbook #ดึงข้อมูลจากไฟล์ excel

def open_window_4():
    def showInfo():
        x = np.linspace(1, 10, 10)
        plt.plot(x, x)
        plt.show()

    def showSumInfo():
        x = np.linspace(1, 10, 10)
        plt.plot(x, x)
        plt.show()

    window4 = Tk()
    window4.geometry("900x700")
    window4.configure(bg="#FFFFFF")

    canvas = Canvas(window4, bg="#FFFFFF", height=700, width=900, bd=0, highlightthickness=0, relief="ridge")
    canvas.place(x=0, y=0)
    canvas.create_rectangle(-1.0, 1011.0, 1920.0093994140625, 1012.0, fill="#000000", outline="")

    subjectID = IntVar()
    canvas.create_text(20.0, 20.0, anchor="nw", text="โปรดใส่รหัสวิชาที่ต้องการ", fill="#000000", font=("Inter SemiBold", 15 * -1))
    entry_s1 = ttk.Entry(window4, style="TEntry", textvariable=subjectID)
    entry_s1.place(x=20.0, y=60.0, width=165.0, height=26.0)

    choiceInCome = StringVar(window4, value="โปรดเลือก")
    combo1 = ttk.Combobox(window4, width=20, textvariable=choiceInCome)
    combo1["value"] = ("งบประมาณแผ่นดิน", "งบประมาณรายได้")
    combo1.place(x=20.0, y=100.0)

    choiceSum = StringVar(window4, value="โปรดเลือก")
    combo2 = ttk.Combobox(window4, width=20, textvariable=choiceSum)
    combo2["value"] = ("รายเดือน", "รายไตรมาส", "รายปี")
    combo2.place(x=20.0, y=140.0)

    buttonA = Button(window4, text=" ตกลง ", font=("Inter SemiBold", 15 * -1), command=showInfo)
    buttonA.place(x=20, y=180)
    
    buttonB = Button(window4, text=" แสดงผลรวมงบประมาณทั้งหมด", font=("Inter SemiBold", 15 * -1), command=showSumInfo)
    buttonB.place(x=20, y=620)

    window4.mainloop()

