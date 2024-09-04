
from pathlib import Path
from tkinter import Tk, Canvas, Entry, Text, Button, PhotoImage
import tkinter.messagebox
from tkinter import *  
from tkinter import ttk

OUTPUT_PATH = Path(__file__).parent
ASSETS_PATH = OUTPUT_PATH / Path(r"E:\Desktop\ฝึกงาน\cp\cp\assets\frame0")

def relative_to_assets(path: str) -> Path:
    return ASSETS_PATH / Path(path)

#จอ1
window_1 = Tk()

window_1.geometry("266x183")
window_1.configure(bg = "#FFFFFF")

#กําหนดขอบ
style = ttk.Style()
style.configure("TEntry", padding=5, relief="flat", background="#EEEEEE")
style.map("TEntry", fieldbackground=[("active", "#FFFFFF"), ("!active", "#C0C0C0")])

canvas = Canvas(window_1,bg = "#FFFFFF",height = 183,width = 266,bd = 0,highlightthickness = 0,relief = "ridge")

canvas.place(x = 0, y = 0)
canvas.create_text(98.0,48.0,anchor="nw",text="ใส่รหัสรายวิชา",fill="#000000",font=("Inter", 12 * -1))

txt = IntVar()
entry_image_1 = PhotoImage(file=relative_to_assets("entry_1.png"))
entry_bg_1 = canvas.create_image(133.5,86.5,image=entry_image_1)
entry_1 = Entry(window_1,textvariable=txt,font = 20,bd=0,bg="#D9D9D9",fg="#000716",highlightthickness=0)
entry_1.place(x=63.0,y=73.0,width=141.0,height=25.0)

def openwindow_2():
    if txt.get() == 1:
        window2 = Tk()

        window2.geometry("900x760")
        window2.configure(bg = "#FFFFFF")

        canvas = Canvas(window2,bg = "#FFFFFF",height = 760,width = 900,bd = 0,highlightthickness = 0,relief = "ridge")

        canvas.place(x = 0, y = 0)
        canvas.create_rectangle(-1.0,1011.0,1920.0093994140625,1012.0,fill="#000000",outline="")
        
        
        x3 = StringVar()
        x4 = StringVar()

        x2,x5,x6,x7,x8,x9,x10 = IntVar(),IntVar(),IntVar(),IntVar(),IntVar(),IntVar(),IntVar()
        x11,x12,x13,x14,x15,x16 = IntVar(),IntVar(),IntVar(),IntVar(),IntVar(),IntVar()

        #ชื่อหัว
        canvas.create_rectangle(9.0,21.0,296.0,58.0,fill="#EEEEEE",outline="")
        canvas.create_text(22.0,31.0,anchor="nw",text="501321 การพยาบาลเด็ก",fill="#000000",font=("Inter SemiBold", 15 * -1))
        canvas.create_rectangle(-0.998779296875,80.95834350585938,900.00341796875,81.95834350585938,fill="#000000",outline="")

        #กล่องป้อนข้อมูล วัน เดือน ปี
        canvas.create_text(17.0,112.0,anchor="nw",text="วัน/เดือน/ปี",fill="#000000",font=("Inter SemiBold", 15 * -1))
        entry_2 = ttk.Entry(window2,style="TEntry",textvariable = x2)
        entry_2.place(x=131.0,y=105.0,width=165.0,height=26.0)

        #กล่องป้อนข้อมูล อ้างอิง
        canvas.create_text(17.0,160.0,anchor="nw",text="อ้างอิง",fill="#000000",font=("Inter SemiBold", 15 * -1))
        entry_3 = ttk.Entry(window2,style="TEntry",textvariable = x3)
        entry_3.place(x=131.0,y=154.0,width=165.0,height=26.0)

        #กล่องป้อนข้อมูล รายการชื่อ-สกุล
        canvas.create_text(17.0,209.0,anchor="nw",text="รายการ/ชื่อ-สกุล",fill="#000000",font=("Inter SemiBold", 15 * -1))
        entry_4 = ttk.Entry(window2,style="TEntry",textvariable = x4)
        entry_4.place(x=131.0,y=203.0,width=165.0,height=25.0)

        #รายการเบิก ค่าตอบแทน
        canvas.create_text(17.0,256.0,anchor="nw",text="รายการเบิก",fill="#000000",font=("Inter SemiBold", 18 * -1))
        canvas.create_text(131.0,262.0,anchor="nw",text="ค่าตอบแทน",fill="#000000",font=("Inter SemiBold", 15 * -1))
        #กล่องป้อนข้อมูล ค่าสอน
        entry_5 = ttk.Entry(window2,style="TEntry",textvariable = x5)
        entry_5.place(x=131.0,y=286.0,width=165.0,height=26.0)
        canvas.create_text(47.0,296.0,anchor="nw",text="ค่าสอน",fill="#000000",font=("Inter SemiBold", 15 * -1))
        #กล่องป้อนข้อมูล แหล่งฝึก
        entry_6 = ttk.Entry(window2,style="TEntry",textvariable = x6)
        entry_6.place(x=131.0,y=335.0,width=165.0,height=26.0)
        canvas.create_text(47.0,345.0,anchor="nw",text="แหล่งฝึก",fill="#000000",font=("Inter SemiBold", 15 * -1))
        #กล่องป้อนข้อมูล ปฐมนิเทศ
        entry_7 = ttk.Entry(window2,style="TEntry",textvariable = x7)
        entry_7.place(x=131.0,y=389.0,width=165.0,height=26.0)
        canvas.create_text(47.0,399.0,anchor="nw",text="ปฐมนิเทศ",fill="#000000",font=("Inter SemiBold", 15 * -1))

        #ค่าใช้สอย
        canvas.create_text(131.0,433.0,anchor="nw",text="ค่าใช้สอย",fill="#000000",font=("Inter SemiBold", 15 * -1))
        #กล่องป้อนข้อมูล ค่าจ้างเหมารถ
        entry_8 = ttk.Entry(window2,style="TEntry",textvariable = x8)
        entry_8.place(x=131.0,y=452.0,width=165.0,height=25.0)
        canvas.create_text(23.0,461.0,anchor="nw",text="ค่าจ้างเหมารถ",fill="#000000",font=("Inter SemiBold", 15 * -1))

        #ค่าวัสดุ
        canvas.create_text(131.0,501.0,anchor="nw",text="ค่าวัสดุ",fill="#000000",font=("Inter SemiBold", 15 * -1))
        #กล่องป้อนข้อมูล นํ้ามันเชื้อเพลิง
        entry_9 = ttk.Entry(window2,style="TEntry",textvariable = x9)
        entry_9.place(x=131.0,y=520.0,width=165.0,height=25.0)
        canvas.create_text(23.0,529.0,anchor="nw",text="นํ้ามันเชื้อเพลิง",fill="#000000",font=("Inter SemiBold", 15 * -1))

        #รายการเบิก - คณาอาจารย์ และ ใช้สอย
        canvas.create_text(386.0,256.0,anchor="nw",text="รายการเบิก-คณาอาจารย์",fill="#000000",font=("Inter SemiBold", 18 * -1))
        canvas.create_text(608.0,265.0,anchor="nw",text="ใช้สอย",fill="#000000",font=("Inter SemiBold", 15 * -1))
        #กล่องป้อนข้อมูล ค่าเบี้ยเลี้ยง
        entry_10 = ttk.Entry(window2,style="TEntry",textvariable = x10)
        entry_10.place(x=608.0,y=289.0,width=164.0,height=26.0)
        canvas.create_text(514.0,297.0,anchor="nw",text="ค่าเบี้ยเลี้ยง",fill="#000000",font=("Inter SemiBold", 15 * -1))
        #กล่องป้อนข้อมูล ค่าที่พัก
        entry_11 = ttk.Entry(window2,style="TEntry",textvariable = x11)
        entry_11.place(x=608.0,y=338.0,width=164.0,height=26.0)
        canvas.create_text(514.0,344.0,anchor="nw",text="ค่าที่พัก",fill="#000000",font=("Inter SemiBold", 15 * -1))
        #กล่องป้อนข้อมูล ค่าเดินทาง
        entry_12 = ttk.Entry(window2,style="TEntry",textvariable = x12)
        entry_12.place(x=608.0,y=392.0,width=164.0,height=26.0)
        canvas.create_text(514.0,400.0,anchor="nw",text="ค่าเดินทาง",fill="#000000",font=("Inter SemiBold", 15 * -1))

        #ข้อความ รายการเบิกนิสิต และ ใช้สอย
        canvas.create_text(447.0,434.0,anchor="nw",text="รายการเบิก-นิสิต",fill="#000000",font=("Inter SemiBold", 18 * -1))
        canvas.create_text(608.0,443.0,anchor="nw",text="ใช้สอย",fill="#000000",font=("Inter SemiBold", 15 * -1))
        #กล่องป้อนข้อมูล ค่าที่พัก
        entry_13 = ttk.Entry(window2,style="TEntry",textvariable = x13)
        entry_13.place(x=608.0,y=466.0,width=164.0,height=25.0)
        canvas.create_text(514.0,474.0,anchor="nw",text="ค่าที่พัก",fill="#000000",font=("Inter SemiBold", 15 * -1))
        #กล่องป้อนข้อมูล ค่าเดินทาง
        entry_14 = ttk.Entry(window2,style="TEntry",textvariable = x14)
        entry_14.place(x=608.0,y=514.0,width=164.0,height=26.0)
        canvas.create_text(514.0,523.0,anchor="nw",text="ค่าเดินทาง",fill="#000000",font=("Inter SemiBold", 15 * -1))

        #ข้อความ รายการเบิก-พนักงานขับรถ และ ใช้สอย
        canvas.create_text(368.0,560.0,anchor="nw",text="รายการเบิก-พนักงานขักรถ",fill="#000000",font=("Inter SemiBold", 18 * -1))
        canvas.create_text(608.0,569.0,anchor="nw",text="ใช้สอย",fill="#000000",font=("Inter SemiBold", 15 * -1))
        #กล่องป้อนข้อมูล เบี้ยเลี้ยง
        entry_15 = ttk.Entry(window2,style="TEntry",textvariable = x15)
        entry_15.place(x=608.0,y=593.0,width=164.0,height=25.0)
        canvas.create_text(521.0,601.0,anchor="nw",text="เบี้ยเลี้ยง",fill="#000000",font=("Inter SemiBold", 15 * -1))
        #กล่องป้อนข้อมูล OT
        entry_16=ttk.Entry(window2,style="TEntry",textvariable = x16)
        entry_16.place(x=608.0,y=642.0,width=164.0,height=25.0)
        canvas.create_text(528.0,649.0,anchor="nw",text="OT",fill="#000000",font=("Inter SemiBold", 15 * -1))

        def success():
            confirm = tkinter.messagebox.askquestion("ยืนยัน","คุณต้องส่งข้อมูลหรือไม่" )
            if confirm == "yes" :
                tkinter.messagebox.showinfo("Success","ส่งข้อมูลเรียบร้อย")
                print(entry_2.get())
                sum = int(entry_5.get()) + int(entry_6.get()) + int(entry_7.get()) + int(entry_8.get()) + int(entry_9.get()) + int(entry_10.get()) + int(entry_11.get()) + int(entry_12.get()) + int(entry_13.get()) + int(entry_14.get()) + int(entry_15.get()) + int(entry_16.get())
                print(sum)
                window2.destroy()

        def deletee():
            tkinter.messagebox.showinfo("Delete","ล้างข้อมูลเรียบร้อย")
            entry_2.delete(0,END)
            entry_3.delete(0,END)
            entry_4.delete(0,END)
            entry_5.delete(0,END)
            entry_6.delete(0,END)
            entry_7.delete(0,END)
            entry_8.delete(0,END)
            entry_9.delete(0,END)
            entry_10.delete(0,END)
            entry_11.delete(0,END)
            entry_12.delete(0,END)
            entry_13.delete(0,END)
            entry_14.delete(0,END)
            entry_15.delete(0,END)
            entry_16.delete(0,END)
        
        money = IntVar()
        Radiobutton(window2,text = "งบประมาณแผ่นดิน",font=("Inter SemiBold", 15 * -1),variable = money,value = 1 ).place(x=100,y=620)
        Radiobutton(window2,text = "งบประมาณรัฐบาล",font=("Inter SemiBold", 15 * -1),variable = money,value = 2 ).place(x=300,y=620)
        
        button1 = Button(window2,text = " ส่งข้อมูล ",font=("Inter SemiBold", 15 * -1),command = success).place(x=100,y=680)
        button2 = Button(window2,text = " ล้างข้อมูล ",font=("Inter SemiBold", 15 * -1),command = deletee).place(x=300,y=680)
    else : 
        tkinter.messagebox.showinfo("INVALID","โปรดใส่รหัสวิชาให้ถูกต้อง")

button_image_1 = PhotoImage(file=relative_to_assets("button_1.png"))
button_1 = Button(image=button_image_1,borderwidth=0,highlightthickness=0,relief="flat",command = openwindow_2)
button_1.place(x=105.0,y=105.0,width=50.0,height=23.0)

window_1.resizable(False, False)
window_1.mainloop()
