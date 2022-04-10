from tkinter import *
import openpyxl


window = Tk()
window.geometry("550x400")
title = Label(window , text = "Excel Python" , font = ("Lucida Calligraphy" , 20)).place(relx = 0.5 , rely = 0.15 , anchor = "center")
def mark():
    wb = openpyxl.Workbook()
    sheet = wb.active
    fn = filename.get()
    cr = cellrow.get()
    cc = cellcolumn.get()
    cv = cellvalue.get()
    ccccr = cc+cr
    c1 = sheet[ccccr]
    c1.value = cv
    wb.save(fn + ".xlsx")
filename = StringVar()
cellrow = StringVar()
cellcolumn = StringVar()
cellvalue = StringVar()

l1 = Label(window , text = "File name").place(relx = 0.07 , rely = 0.3 , anchor = "center")
l2 = Label(window , text = "Cell name").place(relx = 0.07 , rely = 0.45 , anchor = "center")
l3 = Label(window , text = "Cell value").place(relx = 0.07 , rely = 0.6 , anchor = "center")

e1 = Entry(window , textvar = filename).place(relx = 0.3 , rely = 0.3 , anchor = "center")
e2 = Entry(window , textvar = cellcolumn).place(relx = 0.3 , rely = 0.45 , anchor = "center")
e3 = Entry(window , textvar = cellrow).place(relx = 0.7 , rely = 0.45 , anchor = "center")
e4 = Entry(window , textvar = cellvalue).place(relx = 0.3 , rely = 0.6 , anchor = "center")

b1 = Button(window , text = "Mark" , command = mark).place(relx = 0.5 , rely = 0.88 , anchor = "center")
window.mainloop()