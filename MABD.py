import tkinter as tkr
import tkinter.messagebox as tkrmsg
import pandas as pd
import numpy as np
from copy import deepcopy


def main(db_name):
    global data, root, Database_Size, Database_Name
    Database_Name = db_name
    data = pd.read_excel(Database_Name)
    Database_Size = len(data)
    save_database()

    root = tkr.Tk()

    showall_button = tkr.Button(root, text='عرض كل الأصناف', font=('Arial', '20'))
    add_button = tkr.Button(root, text='إضافة صنف', font=('Arial', '20'))
    edit_button = tkr.Button(root, text='تعديل صنف', font=('Arial', '20'))
    report_button = tkr.Button(root, text='بدء فاتورة', font=('Arial', '20'))

    showall_button.bind("<Button-1>", showall)
    add_button.bind("<Button-1>", addone)
    edit_button.bind("<Button-1>", temp)
    report_button.bind("<Button-1>", temp)

    showall_button.pack(side=tkr.LEFT)
    add_button.pack(side=tkr.LEFT)
    edit_button.pack(side=tkr.LEFT)
    report_button.pack(side=tkr.LEFT)

    root.mainloop()


def save_database():
    for i in range(Database_Size):
        mi = i
        for j in range(i+1, Database_Size):
            if data['Name'][j] < data['Name'][mi]:
                mi = j
        x = deepcopy(data.loc[mi])
        y = deepcopy(data.loc[i])
        data.loc[mi] = deepcopy(y)
        data.loc[i] = deepcopy(x)
    writer = pd.ExcelWriter(Database_Name, engine='xlsxwriter')
    data.to_excel(writer, sheet_name='Sheet1')
    writer.save()


def get_index(name):
    found = 0
    for i in data["Name"]:
        if i == name:
            return found
        found += 1
    return found


def notnumber(stri):
    x = 0
    dots = 0

    for i in stri:
        if i == '.':
            dots += 1
        elif i > '9' or i < '0':
            x = 1
            break
    if dots > 1 or x > 0:
        return True
    else:
        return False


def showall(event):
    showall_window = tkr.Tk()
    cpr_label = tkr.Label(showall_window, text='  سعر المستهلك  ', font=('Arial', '25'))
    name_label = tkr.Label(showall_window, text='  أسم الصنف  ', font=('Arial', '25'))
    phpr_label = tkr.Label(showall_window, text='  سعر الصيدلي ', font=('Arial', '25'))

    cpr_label.grid(row=0, column=0, columnspan=2)
    name_label.grid(row=0, column=2, columnspan=4)
    phpr_label.grid(row=0, column=6, columnspan=2)
    for i in range(Database_Size):
        cpr_label = tkr.Label(showall_window, text=("%.2f" % data['CPrice'][i]), font=('Arial', '15'))
        name_label = tkr.Label(showall_window, text=data['Name'][i], font=('Arial', '15'))
        phpr_label = tkr.Label(showall_window, text=("%.2f" % data['PhPrice'][i]), font=('Arial', '15'))

        cpr_label.grid(row=i+1, column=0, columnspan=2)
        name_label.grid(row=i+1, column=2, columnspan=4)
        phpr_label.grid(row=i+1, column=6, columnspan=2)

    showall_window.mainloop()


def addone(event):
    global cpr_entry_add, name_entry_add, phpr_entry_add
    addone_window = tkr.Tk()

    cpr_label = tkr.Label(addone_window, text='سعر المستهلك', font=('Arial', '20'))
    name_label = tkr.Label(addone_window, text='أسم الصنف', font=('Arial', '20'))
    phpr_label = tkr.Label(addone_window, text='سعر الصيدلي', font=('Arial', '20'))

    cpr_entry_add = tkr.Entry(addone_window, font=('Arial', '20'), width=40)
    name_entry_add = tkr.Entry(addone_window, font=('Arial', '20'), width=40)
    phpr_entry_add = tkr.Entry(addone_window, font=('Arial', '20'), width=40)

    add_button = tkr.Button(addone_window, text='إضافة', font=('Arial', '20'), width=10)
    add_button.bind("<Button-1>", add_button_pressed)

    cpr_label.grid(row=0)
    name_label.grid(row=1)
    phpr_label.grid(row=2)

    cpr_entry_add.grid(row=0, column=1)
    name_entry_add.grid(row=1, column=1)
    phpr_entry_add.grid(row=2, column=1)

    add_button.grid(row=3, columnspan=2)

    addone_window.mainloop()


def add_button_pressed(event):
    if not cpr_entry_add.get() or not name_entry_add.get() or not phpr_entry_add.get():
        tkrmsg.showerror('خطأ', 'لا يمكن ترك احد الخانات فارغة')

    elif get_index(name_entry_add.get()) < Database_Size:
        tkrmsg.showerror('خطأ', 'هذا الأسم موجود من قبل')

    else:
        if notnumber(cpr_entry_add.get()):
            tkrmsg.showerror('خطأ', 'خطأ فى سعر المستهلك')
        elif notnumber(phpr_entry_add.get()):
            tkrmsg.showerror('خطأ', 'خطأ فى سعر الصيدلي')
        else:
            add_product(name_entry_add.get(), float('0' + cpr_entry_add.get() + '0'), float('0' + phpr_entry_add.get() + '0'))
            tkrmsg.showinfo('تم', 'تم اضافة الصنف بنجاح')


def add_product(name, cpr, phpr):
    global Database_Size
    data.loc[Database_Size] = [name, cpr, phpr]
    Database_Size += 1
    save_database()


def search(name, cpr, phpr):
    sz = 0
    df = DataFrame(columns=(i for i in data))
    for i in range(Database_Size):
        t = data.loc[i]
        if ((not name) or (name in t["Name"])) and ((not cpr) or (cpr == t["CPrice"])) and ((not phpr) or (phpr == t["PhPrice"])):
            df.loc[sz] = data.loc[i]
            sz += 1
    return df


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


main("Book.xlsx")
