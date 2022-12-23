from tkinter import *
from tkinter import ttk

from PIL import Image, ImageTk


class SummaryTab:

    def __init__(self, tab):
        """Constructor"""
        self.tab = tab
        image2 = Image.open(".\img\\20200106_115612.png")
        image2 = image2.resize((700, 400), Image.ANTIALIAS)
        background_image2 = ImageTk.PhotoImage(image2)
        background_label2 = Label(tab, image=background_image2)
        background_label2.place(x=0, y=0, relwidth=1, relheight=1)
        print("Hi")
