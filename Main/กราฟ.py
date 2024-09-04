from tkinter import *
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

def open_window_4():
    def showInfo():
        x = np.linspace(1, 10, 10)
        plt.clf()  # Clear previous plot
        plt.plot(x, x)
        canvas.draw()

    def showSumInfo():
        x = np.linspace(1, 10, 10)
        plt.clf()  # Clear previous plot
        plt.plot(x, x)
        canvas.draw()

    window4 = Tk()
    window4.geometry("900x700")
    window4.configure(bg="#FFFFFF")

    # Create a Matplotlib figure and embed it in the Tkinter window
    fig, ax = plt.subplots()
    canvas = FigureCanvasTkAgg(fig, master=window4)
    canvas_widget = canvas.get_tk_widget()
    canvas_widget.pack(side=TOP, fill=BOTH, expand=1)

    # Call the showInfo function by default
    showInfo()

    window4.mainloop()

# Call the function to open the window
open_window_4()
