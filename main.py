import tkinter as tk
from tkinter import END, ttk
import tkinter
from tkcalendar import DateEntry
from tkinter import messagebox
import os
import openpyxl
from openpyxl import Workbook

#Fix calender and icon

def clear_input():

    msg_box = tk.messagebox.askquestion('DELETE ALL', 'Are you sure you want to clear the form?',
                                        icon='warning')
    if msg_box == 'yes':                                    
        for widget in user_info_frame.winfo_children():
            if isinstance(widget, tk.Entry):
                widget.delete(0, tk.END)
        for widget in stringing_frame.winfo_children():
            if isinstance(widget, tk.Entry):
                widget.delete(0, tk.END)
        for widget in other_info_frame.winfo_children():
            if isinstance(widget, tk.Entry):
                widget.delete(0, tk.END)
        for widget in cost_frame.winfo_children():
            if isinstance(widget, tk.Entry):
                widget.delete(0, tk.END)
            if isinstance(widget,tk.Checkbutton):
                widget.deselect()


def enter_data():

    #User info
    firstname = first_name_entry.get()
    lastname = last_name_entry.get()
    useremail = email_entry.get()
    phone_number = number_entry.get()

    # Racquet Info
    racquet = racket_entry.get()
    racquet_model = Model_entry.get()
    string_pat = pattern_entry.get()
    skipH = skip_H_entry.get()
    skipT = skip_T_entry.get()
    tieH = tie_H_entry.get()
    tieT = tie_T_entry.get()
    ten = tension_entry.get()

    # Stringing Info
    stringType = string_combobox.get()
    stringBrand = string_brand_entry.get()
    stringModel = string_model_entry.get()
    sten = stencil_combobox.get()
        
    # Expenses
    lab = labour_status.get()
    stringCost = string_cost_entry.get()
    stenCost = sten_status.get()
    calender = cal.get()

    if stringCost == "": # So the program doesnt break if String Cost is empty lol ¯\_(ツ)_/¯
        stringCost = 0

    # Cost with or without certain feilds
    total_cost = (float(lab) + float(stenCost) + float(stringCost))

    empty = ""

    filepath = "racquet_stringing.xlsx"

    if not os.path.exists(filepath):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        heading = ["First Name", "Last Name", "Email", "Phone Number", " ",
                    "Racquet", "Model", "String Pattern", "Skip H", "Skip T", "Tie H", "Tie T", "Tension", " ",
                     "String Type", "String Brand", "String Model", "Stencil", " ",
                      "Date","Total Cost"]
        sheet.append(heading)
        workbook.save(filepath)
    workbook = openpyxl.load_workbook(filepath)
    sheet = workbook.active
    sheet.append([firstname,lastname,useremail,phone_number,empty, 
                    racquet,racquet_model,string_pat,skipH,skipT,tieH,tieT,ten,empty,
                    stringType,stringBrand,stringModel,sten,empty,
                    calender,"${:.2f}".format(total_cost)])
    workbook.save(filepath)



window = tk.Tk()
window.title("Racquet Stringing Form")

frame = tk.Frame(window)
frame.pack()

# Entering users info
user_info_frame = tk.LabelFrame(frame, text="User Information")
user_info_frame.grid(row=0, column=0,padx=20,pady=10)


first_name_label = tk.Label(user_info_frame, text="First Name")
first_name_label.grid(row=0,column=0)

last_name_label = tk.Label(user_info_frame, text="Last Name")
last_name_label.grid(row=0,column=1)

first_name_entry = tk.Entry(user_info_frame)
last_name_entry = tk.Entry(user_info_frame)
first_name_entry.grid(row=1,column=0)
last_name_entry.grid(row=1,column=1)


email_label = tk.Label(user_info_frame,text="Email")
email_label.grid(row=2,column=0)
email_entry = tk.Entry(user_info_frame)
email_entry.grid(row=3,column=0)

number_label = tk.Label(user_info_frame,text="Phone Number")
number_label.grid(row=2,column=1)
number_entry = tk.Entry(user_info_frame)
number_entry.grid(row=3,column=1)

for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=10,pady=5)

# Stringing Infomation

stringing_frame = tk.LabelFrame(frame)
stringing_frame.grid(row=1,column=0,sticky="news",padx=20,pady=10)

string_label = tk.Label(stringing_frame, text="Racquet Information")
string_label.grid(row=0,column=0)

racket_Brand = tk.Label(stringing_frame, text="Racquet Brand")
racket_Brand.grid(row=1,column=0)
racket_entry = tk.Entry(stringing_frame)
racket_entry.grid(row=2,column=0)

racket_Model = tk.Label(stringing_frame, text="Racquet Model")
racket_Model.grid(row=1,column=1)
Model_entry = tk.Entry(stringing_frame)
Model_entry.grid(row=2,column=1)

string_pattern = tk.Label(stringing_frame, text="String Pattern")
string_pattern.grid(row=1,column=2)
pattern_entry = tk.Entry(stringing_frame)
pattern_entry.grid(row=2,column=2)

skip_H = tk.Label(stringing_frame, text="Skip Head")
skip_H.grid(row=1,column=3)
skip_H_entry = tk.Entry(stringing_frame)
skip_H_entry.grid(row=2,column=3)

skip_T = tk.Label(stringing_frame, text="Skip Throat")
skip_T.grid(row=3,column=0)
skip_T_entry = tk.Entry(stringing_frame)
skip_T_entry.grid(row=4,column=0)

tie_H = tk.Label(stringing_frame, text="Tie Head")
tie_H.grid(row=3,column=1)
tie_H_entry = tk.Entry(stringing_frame)
tie_H_entry.grid(row=4,column=1)

tie_T = tk.Label(stringing_frame, text="Tie Throat")
tie_T.grid(row=3,column=2)
tie_T_entry = tk.Entry(stringing_frame)
tie_T_entry.grid(row=4,column=2)

tension_type = tk.Label(stringing_frame, text="Tension")
tension_type.grid(row=3,column=3)
tension_entry = tk.Entry(stringing_frame)
tension_entry.grid(row=4,column=3)

for widget in stringing_frame.winfo_children():
    widget.grid_configure(padx=10,pady=5)

# Stringing Info Frame

other_info_frame = tk.LabelFrame(frame)
other_info_frame.grid(row=2,column=0,sticky="news",padx=20,pady=10)
other_label = tk.Label(other_info_frame, text="String Infomation")
other_label.grid(row=0,column=0)

string_type = tk.Label(other_info_frame, text="String Type")
string_type.grid(row=1,column=0)
string_combobox = ttk.Combobox(other_info_frame, values=["Synthetic Gut", "Multi-Filament","Natural Gut","Polyester", "Other"])
string_combobox.grid(row=2,column=0)

string_brand = tk.Label(other_info_frame, text="String Brand")
string_brand.grid(row=1,column=1)
string_brand_entry = tk.Entry(other_info_frame)
string_brand_entry.grid(row=2,column=1)

string_model = tk.Label(other_info_frame, text="String Model")
string_model.grid(row=1,column=2)
string_model_entry = tk.Entry(other_info_frame)
string_model_entry.grid(row=2,column=2)

stencil_option = tk.Label(other_info_frame, text='Stencil')
stencil_combobox = ttk.Combobox(other_info_frame, values=["None","Head", "Prince", "Babolat", "Dunlop", "Yonex","Wilson"])
stencil_option.grid(row=3,column=1)
stencil_combobox.grid(row=4,column=1)

for widget in other_info_frame.winfo_children():
    widget.grid_configure(padx=25,pady=5)


# Expense Frame

cost_frame = tk.LabelFrame(frame)
cost_frame.grid(row=3,column=0,sticky="news",padx=20,pady=10)
expense_label = tk.Label(cost_frame, text="Expense Information")
expense_label.grid(row=0,column=0)

LABOUR_COST = 25  # Default Labour cost
labour = tk.Label(cost_frame,text="Labour: ${}".format(LABOUR_COST))
labour.grid(row=1,column=0)
labour_status = tkinter.StringVar(value=0)
labour_entry = tk.Checkbutton(cost_frame,variable=labour_status,onvalue=LABOUR_COST,offvalue=0)
labour_entry.grid(row=2,column=0)

string_cost = tk.Label(cost_frame, text="String Cost")
string_cost.grid(row=1,column=2)
string_cost_entry = tk.Entry(cost_frame)
string_cost_entry.insert(0,"0")
string_cost_entry.grid(row=2,column=2)

STENCIL_COST = 2  # Default Stencil cost
stencil_cost = tk.Label(cost_frame,text="Stencil Cost: ${}".format(STENCIL_COST))
stencil_cost.grid(row=1,column=3)
sten_status = tkinter.StringVar(value=0)
stencil_cost_entry = tk.Checkbutton(cost_frame,variable=sten_status,onvalue=STENCIL_COST,offvalue=0)
stencil_cost_entry.grid(row=2,column=3)

date = tk.Label(cost_frame,text="Date")
date.grid(row=1,column=5)

cal = DateEntry(cost_frame,selectmode='day')
cal.grid(row=2,column=5)

for widget in cost_frame.winfo_children():
    widget.grid_configure(padx=15,pady=5)

# Buttons
 
enter_button = tk.Button(frame, text="Enter Data", command=enter_data, bg="green",font="bold")
enter_button.grid(row=4,column=0,sticky="news")

clear_button = tk.Button(frame, text="Clear All", command=clear_input, bg="indian red",font="bold")
clear_button.grid(row=5,column=0,sticky="news",pady=10)


window.mainloop()