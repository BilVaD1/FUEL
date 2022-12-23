import os
import time
from tkinter import *
from tkinter import ttk
from components.summary_tab import SummaryTab
from PIL import Image, ImageTk

from tasks import Tasks

while 1:
    os.system("TASKKILL /F /IM Excel.exe")  # Stop Excel if it is running
    time.sleep(0)
    break

# Dropdown menu options
options = [
    "10",
    "9",
    "12",
    "13",
    "14",
    "21",
    "22",
    "23",
    "7",
    "6",
    "5",
    "17",
    "Overflow",
    "Sludge",
    "Bilge",
    "Waste",
    "MEoil"
]

tanks_for_colors = [] # This is variable for realize the changing of color on each curve of the graphic, it keeps only unique values


'''Create functions that will return values from Entry fields'''


def return_tank_name():
    content = clicked.get()
    return content


def return_sounding():
    content = Sounding_entry.get()
    return content

def return_density():
    content = Density_entry.get()
    return content

def return_temperature():
    content = Temperature_entry.get()
    return content


root = Tk()
root.title('Tanks')
root.geometry('700x400+700+300')
root.resizable(False, False)
#root.config(bg=r'.\\img\\20200106_115612.jpg')
root.iconbitmap(r".\\img\ship_14716.ico")


# Create the Notebook widget
nb = ttk.Notebook(root)

# Pack the Notebook widget to make it visible
nb.pack(expand=1, fill='both')

# Create the first tab
tab1 = ttk.Frame(nb)

# Create the second tab
tab2 = ttk.Frame(nb)

# Add the tabs to the Notebook widget
nb.add(tab1, text='Tank')
nb.add(tab2, text='Summary')


image = Image.open(".\img\\20200106_115612.png")
image1 = image.resize((700, 400), Image.ANTIALIAS)
background_image = ImageTk.PhotoImage(image1)
background_label = Label(tab1, image=background_image)
background_label.place(x=0, y=0, relwidth=1, relheight=1)
  

tank_label = Label(tab1, text='Select the tank: ')
tank_label.grid(row=0, column=0, padx=10, pady=10, sticky=W)

# datatype of menu text
clicked = StringVar()
  
# initial menu text
clicked.set( "Tank Name" )
  
# Create Dropdown menu
drop = OptionMenu( tab1 , clicked , *options )
drop.grid(row=0, column=2, columnspan=1, sticky=E + W, pady=35)

Sounding_label = Label(tab1, text='Please write the Sounding (in the format 0.00) = ')
Sounding_label.grid(row=1, column=0, padx=10, pady=10, sticky=W)

Sounding_entry = Entry(tab1)
Sounding_entry.grid(row=1, column=2, columnspan=2, sticky=W + E)

answer_m3 = Label(tab1, text='Results: ')
answer_m3.grid(row=2, column=2, sticky=W)

answer_m3 = Label(tab1, padx=5, text='There will be the Volume in m3')
answer_m3.grid(row=2, column=2, sticky=E, pady=15, padx=60)

Density_label = Label(tab1, text='Please write the density of fuel at 15 celsius:')
Density_label.grid(row=3, column=0, padx=10, pady=30, sticky=W)

Density_entry = Entry(tab1)
Density_entry.grid(row=3, column=2, columnspan=2, sticky=W + E)

Temperature_label = Label(tab1, text='Please write the current temperature of fuel:')
Temperature_label.grid(row=4, column=0, padx=10, pady=10, sticky=W)

Temperature_entry = Entry(tab1)
Temperature_entry.grid(row=4, column=2, columnspan=2, sticky=W + E)

answer_mass = Label(tab1, padx=5, text='There will be the mass of oil in "t"')
answer_mass.grid(row=5, column=2, sticky=W + E, pady=25)


'''Function for calling class Helpers()'''
def task2():
    tasks = Tasks()
    return tasks.task1(return_sounding(), return_tank_name(), return_density(), return_temperature(), answer_mass, answer_m3, tanks_for_colors)


Start = Button(tab1, command=task2, text='Start')
Start.grid(row=2, column=0, padx=10, pady=10, sticky=W + E)


"""Tab2"""
image2 = Image.open(".\img\\20200106_115612.png")
image2 = image2.resize((700, 400), Image.ANTIALIAS)
background_image2 = ImageTk.PhotoImage(image2)
background_label2 = Label(tab2, image=background_image2)
background_label2.place(x=0, y=0, relwidth=1, relheight=1)


root.mainloop()
