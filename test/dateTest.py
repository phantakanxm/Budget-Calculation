import tkinter  as tk 
from tkcalendar import DateEntry
my_w = tk.Tk()
my_w.geometry("340x220")  

cal=DateEntry(my_w,selectmode='day',date_pattern="dd/mm/yyyy")
cal.place(x=131.0,y=105.0,width=165.0,height=26.0)
dt = cal.get_date()
dt3=dt.strftime("%d/%m/%Y")
dt4=dt.strftime("%m")
l = dt3.split('/')
print(type(dt3))
print(dt4)
my_w.mainloop()