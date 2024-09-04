import matplotlib.pyplot as plt
from tkinter import Tk, Canvas, Button
  
from tkinter import *  
from tkinter import ttk

from openpyxl import load_workbook #ดึงข้อมูลจากไฟล์ excel

from Calculate import *

filename = 'งบประมาณรวม.xlsx'
wb = load_workbook(filename)
ws = wb.active

def addlabels(x,y):
	for i in range(len(x)):
		plt.text(i, y[i], y[i], ha = 'center')

def open_window_4():

    def showInfo():
        income = combo1.get()
        subjectID = int(entry_s1.get())
        #print("Press Show Summation info ")
        if combo2.get() == 'รายเดือน':
            print('รายเดือน')
            month = combo3.get()
            year = entry_year.get()
            m = SubjectMCal(filename,month,year,income,subjectID)
            #print(f"Summation = {m}")

            categories = list(m.keys())
            values = list(m.values())

            # Plotting the bar graph
            addlabels(categories, values)
            plt.bar(categories, values)
            plt.xlabel('Categories')
            plt.ylabel('Values')
            plt.title('Monthly Summary Information')
            plt.show()

        elif combo2.get() == 'รายปี':
            print('รายปี')
            y = SubjectYCal(filename,year,income,subjectID)
            #print(f"Summation = {y}")

            categories = list(y.keys())
            values = list(y.values())

            # Plotting the bar graph
            addlabels(categories, values)
            plt.bar(categories, values)
            plt.xlabel('Categories')
            plt.ylabel('Values')
            plt.title('Yearly Summary Information')
            plt.show()
        
        

    def showSumInfo():
        #print("Press Show Summation info ")
        income = combo1.get()
        if combo2.get() == 'รายเดือน':
            print('รายเดือน')
            month = combo3.get()
            year = entry_year.get()
            m = monthCal(filename,month,year,income)
            #print(f"Summation = {m}")

            categories = list(m.keys())
            values = list(m.values())

            # Plotting the bar graph
            addlabels(categories, values)
            plt.bar(categories, values)
            plt.xlabel('Categories')
            plt.ylabel('Values')
            plt.title('Monthly Summary Information')
            plt.show()
        
        elif combo2.get() == 'รายปี':
            print('รายปี')
            y = yearCal(filename,year,income)
            #print(f"Summation = {y}")
            
            categories = list(y.keys())
            values = list(y.values())

            # Plotting the bar graph
            addlabels(categories, values)
            plt.bar(categories, values)
            plt.xlabel('Categories')
            plt.ylabel('Values')
            plt.title('Yearly Summary Information')
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

    choiceInCome = StringVar(window4, value="เลือกงบประมาณ")
    combo1 = ttk.Combobox(window4, width=20, textvariable=choiceInCome)
    combo1["value"] = ("งบประมาณแผ่นดิน", "งบประมาณรายได้")
    combo1.place(x=20.0, y=100.0)

    choiceSum = StringVar(window4, value="โปรดเลือก")
    combo2 = ttk.Combobox(window4, width=20, textvariable=choiceSum)
    combo2["value"] = ("รายเดือน","รายปี")
    combo2.place(x=20.0, y=140.0)
    
    choiceMonth = IntVar(window4, value="เลือกเดือน")
    combo3 = ttk.Combobox(window4, width=20, textvariable=choiceMonth)
    combo3["value"] = ("1","2","3","4","5","6","7","8","9","10","11","12")
    combo3.place(x=20.0, y=180.0)
    
    def temp_text(e):
        entry_year.delete(0,"end")
        
    yearID = IntVar()
    entry_year = ttk.Entry(window4, style="TEntry", textvariable=yearID)
    entry_year.insert(0,"ปี (ค.ศ.)")
    entry_year.place(x=20.0, y=220.0, width=165.0, height=26.0)
    entry_year.bind("<FocusIn>", temp_text)
    
    
    buttonA = Button(window4, text=" ตกลง ", font=("Inter SemiBold", 15 * -1), command=showInfo)
    buttonA.place(x=20, y=260)
    
    buttonB = Button(window4, text=" แสดงผลรวมงบประมาณทั้งหมด", font=("Inter SemiBold", 15 * -1), command=showSumInfo)
    buttonB.place(x=20, y=620)

    window4.mainloop()

