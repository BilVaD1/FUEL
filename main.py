import os
import time
from tkinter import *
from tkinter import ttk
import tkinter as tk
from PIL import Image, ImageTk

from tasks import Tasks
from summary import Summary



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

def stop_excel():
    os.system("TASKKILL /F /IM Excel.exe")  # Stop Excel if it is running
    time.sleep(0)

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

'''Function for calling class Helpers()'''
def task_tab1():
    stop_excel()
    tasks = Tasks(answer_mass, answer_m3, tanks_for_colors)
    return tasks.task1(return_sounding(), return_tank_name(), return_density(), return_temperature())


'''Functions to hide or show rows'''
def show_Mass():
    Density_label.grid(row=3, column=0, padx=60, pady=30, sticky=W)
    Density_entry.grid(row=3, column=2, columnspan=2, sticky=W + E)
    Temperature_label.grid(row=4, column=0, padx=60, pady=10, sticky=W)
    Temperature_entry.grid(row=4, column=2, columnspan=2, sticky=W + E)
    answer_mass.grid(row=5, column=2, sticky=W + E, pady=25)

def hide_Mass():
    Density_label.grid_forget()
    Density_entry.grid_forget()
    Temperature_label.grid_forget()
    Temperature_entry.grid_forget()
    answer_mass.grid_forget()

'''Function to handle the checkbox value'''
def checkClicked():
    if checkbutton_var.get():
        show_Mass()
    else:
        hide_Mass()


root = Tk()
root.title('Tanks')
root.geometry('1000x600+700+300')
root.resizable(False, False)
root.iconbitmap(r".\\img\ship_14716.ico")

'''Create tabs'''
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


"""Tab1"""
image = Image.open(".\img\\20200106_115612.png")
image1 = image.resize((1000, 600), Image.Resampling.LANCZOS)
background_image = ImageTk.PhotoImage(image1)
background_label = Label(tab1, image=background_image)
background_label.place(x=0, y=0, relwidth=1, relheight=1)
  

tank_label = Label(tab1, text='Select the tank: ')
tank_label.grid(row=0, column=0, padx=60, pady=40, sticky=W)

# datatype of menu text
clicked = StringVar()
  
# initial menu text
clicked.set( "Tank Name" )
  
# Create Dropdown menu
drop = OptionMenu( tab1 , clicked , *options )
drop.grid(row=0, column=2, columnspan=1, sticky=E + W, pady=35)

Sounding_label = Label(tab1, text='Please write the Sounding (in the format 0.00) = ')
Sounding_label.grid(row=1, column=0, padx=60, pady=10, sticky=W)

Sounding_entry = Entry(tab1)
Sounding_entry.grid(row=1, column=2, columnspan=2, sticky=W + E)

answer_m3 = Label(tab1, text='Results: ')
answer_m3.grid(row=2, column=2, sticky=W)

answer_m3 = Label(tab1, padx=5, text='There will be the Volume in m3')
answer_m3.grid(row=2, column=2, sticky=E, pady=15, padx=60)

Density_label = Label(tab1, text='Write the density(kg/m3) of fuel at 15 celsius:')

Density_entry = Entry(tab1)

Temperature_label = Label(tab1, text='Write the current temperature(Celsius) of fuel:')

Temperature_entry = Entry(tab1)

answer_mass = Label(tab1, padx=5, text='There will be the mass of oil in "t"')

Start = Button(tab1, command=task_tab1, text='Start')
Start.grid(row=2, column=0, padx=60, pady=10, sticky=W + E)

checkbutton = tk.Checkbutton(tab1, text="Calculate mass", command=checkClicked)
# Create an IntVar to store the state of the checkbox
checkbutton_var = tk.IntVar()

# Set the checkbox to use the IntVar as its variable
checkbutton.config(variable=checkbutton_var)
checkbutton.grid(row=0, column=0, padx=10, pady=10, sticky=E)





"""Tab2"""

'''Functions to hide or show rows'''
def show_Bunkering():
    Density_bunker.grid(row=2, column=3, padx=10, pady=10, sticky=W + N)
    Density_bunker_entry.grid(row=2, column=4, columnspan=2, pady=10, sticky=W + N)
    Temperature_bunker.grid(row=2, column=3, padx=10, pady=110, sticky=W)
    Temperature_bunker_entry.grid(row=2, column=4, columnspan=2, pady=110, sticky=W)
    button_Bunker.grid(row=2, column=3, sticky=W + S, pady=75, padx=10)


def hide_Bunkering():
    Density_bunker.grid_forget()
    Density_bunker_entry.grid_forget()
    Temperature_bunker.grid_forget()
    Temperature_bunker_entry.grid_forget()
    button_Bunker.grid_forget()
    bulk_answer_label.grid_forget()

'''Function to handle the checkbox value'''
def checkClicked_tab2():
    if checkbutton_var_tab2.get():
        show_Bunkering()
    else:
        hide_Bunkering()

def task_tab2():
    Notification_Block['text'] = 'Notification: All ok.'
    Notification_Block['bg'] = 'green'
    Notification_Block.grid(row=8, column=1, padx=20, pady=10, sticky=W + S)
    #Show the labels
    #Show_results() 
    # Get the indices of the selected items
    selected_indices = listbox.curselection()

    # Get the selected options from the options list
    selected_options = [options[i] for i in selected_indices]
    summary = Summary(tanks_capacity, current_capacity, remaining_capacity, answer_tanks_name, Notification_Block, Warning_Block)
    summary.foundSelectedValues(selected_options)
    # Special variable to handle the Warning_Block about the different dates
    dates = summary.foundSelectedValues(selected_options)

    if Notification_Block.cget("bg") != "red" and len(selected_options) > 0:
        if dates[0]:
            Warning_Block.grid(row=9, column=1, padx=20, sticky=W) # Display the warning label
        else:
            Warning_Block.grid_forget() # Remove the warning label
        Show_results()
    elif len(selected_options) == 0:
        Notification_Block['text'] = 'Notification: Choose the tank(s)'
        Notification_Block['bg'] = 'orange'
        reset_results() # Remove the results
        Warning_Block.grid_forget() # Remove the warning label
    else:
        Warning_Block.grid(row=9, column=1, padx=100, sticky=W) # Display the warning label
        reset_results() # Remove the results

    return dates[1]


def Show_results():

    answer_tanks_name.grid(row=4, column=1, padx=100, sticky=W)
    tanks_capacity.grid(row=5, column=1, padx=100, sticky=W)
    current_capacity.grid(row=6, column=1, padx=100, sticky=W)
    remaining_capacity.grid(row=7, column=1, padx=100, sticky=W)

def reset_results():
    answer_tanks_name.grid_forget()
    tanks_capacity.grid_forget()
    current_capacity.grid_forget()
    remaining_capacity.grid_forget()

def open_excel():
    # Generate report
    Summary().generate_report()
    # Open report 
    os.system('start "excel" "reports\Summary_report_2022-12-28.xlsx"')

def caclBunk():
    # Get the indices of the selected items
    selected_indices = listbox.curselection()

    # Get the selected options from the options list
    selected_options = [options[i] for i in selected_indices]

    density_enter = Density_bunker_entry.get()
    temp_enter = Temperature_bunker_entry.get()

    if density_enter == '' or temp_enter == '':
        bulk_answer_label['text'] = f'Please specify the density & temp.'
        bulk_answer_label['bg'] = 'orange'
        bulk_answer_label.grid(row=3, column=3, sticky=E + S, padx=10)
    elif len(selected_options) == 0:
        bulk_answer_label['text'] = f'Choose the tank(s)'
        bulk_answer_label['bg'] = 'orange'
        bulk_answer_label.grid(row=3, column=3, sticky=E + S, padx=10)
    else:
        density = float(density_enter)
        temp = float(temp_enter)
        volume = float(task_tab2())

        tasks = Tasks()
        answer = tasks.mass(density, temp, volume)
        bulk_answer_label['text'] = f'Tank(s) can hold {answer} tons'
        bulk_answer_label['bg'] = 'pink'
        bulk_answer_label.grid(row=3, column=3, sticky=E + S, padx=10)




image2 = Image.open(".\img\\1638890321103.png")
image2 = image2.resize((1000, 600), Image.Resampling.LANCZOS)
background_image2 = ImageTk.PhotoImage(image2)
background_label2 = Label(tab2, image=background_image2)
background_label2.place(x=0, y=0, relwidth=1, relheight=1)

'''Create the Listbox widget'''
listbox = tk.Listbox(tab2, selectmode='multiple')

# Add options to the Listbox
for option in options:
    listbox.insert(tk.END, option)

# Create a button to toggle the visibility of the Listbox
listbox.grid(row=2, column=2, sticky=N, pady=10, padx=20) # Use the show by default

'''Button to hide/show the multiple menu'''
button_hide = tk.Button(tab2, text="Show/Hide Tanks", 
command=lambda: listbox.grid_forget() if listbox.winfo_ismapped() else listbox.grid(row=2, column=2, sticky=N, pady=10, padx=20))
button_hide.grid(row=1, column=2, sticky=N + W, pady=35)

'''Button to calculate the answer block'''
button_summary = tk.Button(tab2, text="Show result", command=task_tab2)
button_summary.grid(row=1, column=1, sticky=N + W, pady=35, padx=100)

'''Answer block with results'''
answer_tanks_name = Label(tab2, text='Tank(s) name(s): ')
tanks_capacity = Label(tab2, text='Tank(s) capacity: ')
current_capacity = Label(tab2, text='Current capacity: ')
remaining_capacity = Label(tab2, text='Remaining capacity: ')

'''Notifications block'''
Notification_Block = Label(tab2, text='Notification: ', bg='yellow')
Warning_Block = Label(tab2, text='Warning: ', bg='yellow')

'''Button to generate report'''
button_report = tk.Button(tab2, text="Generate/Open Report", command=open_excel, bg='green')
button_report.grid(row=2, column=1, sticky=W + N, padx=100)

'''Checkbutton to show/hide the bunkering block'''
checkbutton_tab2 = tk.Checkbutton(tab2, text="Enable Bunker", command=checkClicked_tab2)
# Create an IntVar to store the state of the checkbox
checkbutton_var_tab2 = tk.IntVar()

# Set the checkbox to use the IntVar as its variable
checkbutton_tab2.config(variable=checkbutton_var_tab2)
checkbutton_tab2.grid(row=1, column=3, padx=10, pady=10, sticky=W)

'''Block to calculate the bunkering fuel/oil'''
Density_bunker = Label(tab2, text='Density(kg/m3) of fuel at 15 celsius:')

Density_bunker_entry = Entry(tab2, bg='grey')

Temperature_bunker = Label(tab2, text='Temperature of bunkering fuel:')

Temperature_bunker_entry = Entry(tab2, bg='grey')

button_Bunker = tk.Button(tab2, text="Calc Bunker", bg='purple', command=caclBunk)

bulk_answer_label = Label(tab2, text='')


root.mainloop()
