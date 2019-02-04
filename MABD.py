import tkinter as tkr
import pandas as pd
import numpy as np
from copy import deepcopy


def save_database():
    writer = pd.ExcelWriter(Database_Name, engine='xlsxwriter')
    data.to_excel(writer, sheet_name='Sheet1')
    writer.save()


def get_index(name):
    found = 0
    for i in data["Name"]:
        if i == name:
            return found
        found += 1


def search(name, cpr, phpr):
    sz = 0
    df = DataFrame(columns=(i for i in data))
    for i in range(Database_Size):
        t = data.loc[i]
        if ((not name) or (name in t["Name"])) and ((not cpr) or (cpr == t["CPrice"])) and ((not phpr) or (phpr == t["PhPrice"])):
            df.loc[sz] = data.loc[i]
            sz += 1
    return df


def add_product(name, cpr, phpr):
    global Database_Size
    if get_index(name) < Database_Size:
        print("This name is already in the database")
        return
    data.loc[Database_Size] = [name, cpr, phpr]
    Database_Size += 1
    save_database()
    # print("Product " + name + " added successfully")


def delete_product(name):
    global Database_Size
    found = get_index(name)

    for i in range(found+1, Database_Size):
        data.loc[i-1] = data.loc[i]
    Database_Size -= 1
    data.drop(data.index[Database_Size], inplace=True)
    save_database()


def update_produce(idx, name, cpr, phpr):
    data.loc[idx]["Name"] = name
    data.loc[idx]["CPrice"] = cpr
    data.loc[idx]["PhPrice"] = phpr
    save_database()


def temp(event):
    print('abc')


def main(db_name):
    global data, root, Database_Size, Database_Name
    Database_Name = db_name
    data = pd.read_excel(Database_Name)
    Database_Size = len(data)

    root = tkr.Tk()
    showall_button = tkr.Button(root, text='عرض كل الأصناف')
    edit_button = tkr.Button(root, text='تعديل صنف')
    report_button = tkr.Button(root, text='بدء فاتورة')

    showall_button.bind("<Button-1>", temp)
    edit_button.bind("<Button-1>", temp)
    report_button.bind("<Button-1>", temp)

    showall_button.pack(side=tkr.LEFT)
    edit_button.pack(side=tkr.LEFT)
    report_button.pack(side=tkr.LEFT)
    root.mainloop()

main("Book.xlsx")

"""
frame1 = tkr.Frame(root)
frame1.pack()
frame2 = tkr.Frame(root)
frame2.pack(side=tkr.BOTTOM)
# frame2.pack()

button1 = tkr.Button(frame1, text='Button1', fg='red')
button2 = tkr.Button(frame1, text='Button2', fg='green')
button3 = tkr.Button(frame1, text='Button3', fg='blue')
button4 = tkr.Button(frame2, text='Button4', fg='orange')

button1.pack(side=tkr.LEFT)
button4.pack(side=tkr.LEFT)
button2.pack(side=tkr.LEFT)
button3.pack(side=tkr.LEFT)

"""
