from tkinter import *
from tkinter.ttk import Combobox
from tkinter import Tk, messagebox
from tkcalendar import DateEntry

import openpyxl
import pathlib
from openpyxl import Workbook

root = Tk()
root.title("Data Entry")
root.geometry('900x800+300+200')
root.resizable(False, False)
root.configure(bg="black")

file = pathlib.Path('Backend_data.xlsx')
if file.exists():
    pass
else:
    file = Workbook()
    sheet = file.active
    sheet['A1'] = "Full Name"
    sheet['B1'] = "PhoneNumber"
    sheet['C1'] = "Age"
    sheet['D1'] = "Gender"
    sheet['E1'] = "Occupation"
    sheet['F1'] = "Address"
    sheet['G1'] = "Cause"
    sheet['H1'] = "Past History"
    sheet['I1'] = "Family History"
    sheet['J1'] = "Treatment Date"
    sheet['K1'] = "Treatment and Fees"

    file.save('Backend_data.xlsx')


def submit():
    name = nameValue.get()
    contact = contactValue.get()
    age = AgeValue.get()
    gender = gender_combobox.get()
    occupation = occupationValue.get()
    address = addressEntry.get(1.0, END)
    cause = causeValue.get()
    pastHistory = pastHistoryValue.get()
    familyHistory = familyHistoryValue.get()
    treatmentDate = treatmentDateValue.get()
    treatmentFees = treatmentFeesValue.get()

    if not age.isdigit():
        messagebox.showerror('Error', 'Age must contain only digits.')
        return

    if len(age) > 3:
        messagebox.showerror('Error', 'Age must be up to 3 digits.')
        return

    if len(contact) != 10 or not contact.isdigit():
        messagebox.showerror('Error', 'Contact number must be a 10-digit number.')
        return

    file = openpyxl.load_workbook('Backend_data.xlsx')
    sheet = file.active
    sheet.cell(column=1, row=sheet.max_row + 1, value=name)
    sheet.cell(column=2, row=sheet.max_row, value=contact)
    sheet.cell(column=3, row=sheet.max_row, value=age)
    sheet.cell(column=4, row=sheet.max_row, value=gender)
    sheet.cell(column=5, row=sheet.max_row, value=occupation)
    sheet.cell(column=6, row=sheet.max_row, value=address)
    sheet.cell(column=7, row=sheet.max_row, value=cause)
    sheet.cell(column=8, row=sheet.max_row, value=pastHistory)
    sheet.cell(column=9, row=sheet.max_row, value=familyHistory)
    sheet.cell(column=10, row=sheet.max_row, value=treatmentDate)
    sheet.cell(column=11, row=sheet.max_row, value=treatmentFees)

    file.save(r'Backend_data.xlsx')

    messagebox.showinfo('info', 'detail added!')

    nameValue.set('')
    contactValue.set('')
    AgeValue.set('')
    occupationValue.set('')
    addressEntry.delete(1.0, END)
    causeValue.set('')
    pastHistoryValue.set('')
    familyHistoryValue.set('')
    treatmentDateValue.set('')
    treatmentFeesValue.set('')


def clear():
    nameValue.set('')
    contactValue.set('')
    AgeValue.set('')
    occupationValue.set('')
    addressEntry.delete(1.0, END)
    causeValue.set('')
    pastHistoryValue.set('')
    familyHistoryValue.set('')
    treatmentDateValue.set('')
    treatmentFeesValue.set('')


def display_selected_date(event):
    selected_date = treatmentDateEntry.get_date()
    selected_date_label.config(text=f"Selected Date: {selected_date}")


def on_enter(event):
    event.widget.tk_focusNext().focus()
    return "break"


Label(root, text="Please fill out this Entry form:", font="arial 13", bg="green", fg="#fff").place(x=340, y=20)

Label(root, text='Name', font=23, bg="green", fg="#fff").place(x=50, y=70)
Label(root, text='Contact No', font=23, bg="green", fg="#fff").place(x=50, y=120)
Label(root, text='Age', font=23, bg="green", fg="#fff").place(x=50, y=170)
Label(root, text='Gender', font=23, bg="green", fg="#fff").place(x=390, y=170)
Label(root, text='Occupation', font=23, bg="green", fg="#fff").place(x=50, y=220)
Label(root, text='Address', font=23, bg="green", fg="#fff").place(x=50, y=270)
Label(root, text='Cause', font=23, bg="green", fg="#fff").place(x=50, y=320)
Label(root, text='Past History', font=23, bg="green", fg="#fff").place(x=50, y=370)
Label(root, text='Family History', font=23, bg="green", fg="#fff").place(x=50, y=420)
Label(root, text='Treatment Date', font=23, bg="green", fg="#fff").place(x=50, y=470)

treatmentDateEntry = DateEntry(root, width=30, bd=2, font=20)  
treatmentDateEntry.place(x=200, y=470)
treatmentDateEntry.bind("<<DateEntrySelected>>", display_selected_date)

selected_date_label = Label(root, text="Selected Date:", font=23, bg="green", fg="#fff")
selected_date_label.place(x=200, y=500)
Label(root, text='Treatment', font=23, bg="green", fg="#fff").place(x=50, y=550)

nameValue = StringVar()
contactValue = StringVar()
AgeValue = StringVar()
occupationValue = StringVar()
causeValue = StringVar()
pastHistoryValue = StringVar()
familyHistoryValue = StringVar()
treatmentDateValue = StringVar()
treatmentFeesValue = StringVar()

nameEntry = Entry(root, textvariable=nameValue, width=45, bd=2, font=20)
contactEntry = Entry(root, textvariable=contactValue, width=20, bd=2, font=20)
ageEntry = Spinbox(root, from_=0, to=150, textvariable=AgeValue, width=15, bd=2, font=20)
occupationEntry = Entry(root, textvariable=occupationValue, width=30, bd=2, font=20)
causeEntry = Entry(root, textvariable=causeValue, width=45, bd=2, font=20)
pastHistoryEntry = Entry(root, textvariable=pastHistoryValue, width=45, bd=2, font=20)
familyHistoryEntry = Entry(root, textvariable=familyHistoryValue, width=45, bd=2, font=20)
treatmentFeesEntry = Entry(root, textvariable=treatmentFeesValue, width=45, bd=2, font=20)

gender_combobox = Combobox(root, values=['Male', 'Female'], font='arial 14', state='r', width=14)
gender_combobox.place(x=480, y=170)
gender_combobox.set('Male')

addressEntry = Text(root, width=50, height=2, bd=2)

nameEntry.place(x=200, y=70)
contactEntry.place(x=200, y=120)
ageEntry.place(x=200, y=170)
occupationEntry.place(x=200, y=220)
causeEntry.place(x=200, y=320)
pastHistoryEntry.place(x=200, y=370)
familyHistoryEntry.place(x=200, y=420)
treatmentFeesEntry.place(x=200, y=550)
addressEntry.place(x=200, y=270)

Button(root, text="Submit", bg="#326273", fg="white", width=15, height=2, command=submit).place(x=200, y=620)
Button(root, text="Clear", bg="#326273", fg="white", width=15, height=2, command=clear).place(x=340, y=620)
Button(root, text="Exit", bg="#326273", fg="white", width=15, height=2, command=lambda: root.destroy()).place(x=480, y=620)

# Binding Enter key press event to move focus to next widget
nameEntry.bind('<Return>', on_enter)
contactEntry.bind('<Return>', on_enter)
ageEntry.bind('<Return>', on_enter)
occupationEntry.bind('<Return>', on_enter)
causeEntry.bind('<Return>', on_enter)
pastHistoryEntry.bind('<Return>', on_enter)
familyHistoryEntry.bind('<Return>', on_enter)
treatmentDateEntry.bind('<Return>', on_enter)
treatmentFeesEntry.bind('<Return>', on_enter)

root.mainloop()
