import subprocess
import os
import openpyxl
import tkinter as tkr
import tkinter.messagebox as tkrmsg
import pandas as pd
import numpy
from copy import deepcopy
from datetime import date


def start():
    global data, Database_Size, Database_Name, phars, phars_Size, phars_Db_Name, LETTERS, colchar, fatora_counter
    LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    phars_Db_Name = "Database/phar.xlsx"
    Database_Name = "Database/Book.xlsx"

    colfile = open("Database/column.txt", "r")
    colchar = int(colfile.read())
    colfile.close()

    counterfile = open("Database/counter.txt", "r")
    fatora_counter = int(counterfile.read())
    counterfile.close()

    if colchar < 26:
        tempcols = "B:" + LETTERS[colchar]
    else:
        tempcols = "B:" + LETTERS[(colchar//26)-1] + LETTERS[colchar % 26]

    phars = pd.read_excel(phars_Db_Name, usecols="B")
    data = pd.read_excel(Database_Name, usecols=tempcols)

    phars_Size = len(phars)
    Database_Size = len(data)

    save_database()
    main()


def main():
    global root
    root = tkr.Tk()
    root.title('Mr/Abd Elrahman')
    root.protocol("WM_DELETE_WINDOW", on_closing2)

    showall_button = tkr.Button(root, text='عرض كل الأصناف', font=('Arial', '20'))
    add_button = tkr.Button(root, text='إضافة صنف', font=('Arial', '20'))
    edit_button = tkr.Button(root, text='تعديل صنف', font=('Arial', '20'))
    report_button = tkr.Button(root, text='بدء فاتورة', font=('Arial', '20'))

    showall_button.bind("<Button-1>", showall)
    add_button.bind("<Button-1>", addone)
    edit_button.bind("<Button-1>", editone)
    report_button.bind("<Button-1>", report)

    showall_button.pack(side=tkr.LEFT)
    add_button.pack(side=tkr.LEFT)
    edit_button.pack(side=tkr.LEFT)
    report_button.pack(side=tkr.LEFT)

    root.mainloop()


def canvasfunc(event):
    canvas_rightframe_fatora.configure(scrollregion=canvas_rightframe_fatora.bbox("all"), width=675, height=500)


def canvasfunc2(event):
    canvas_downframe_fatora.configure(scrollregion=canvas_downframe_fatora.bbox("all"), width=725, height=500)


def showall_canvas_func0(event):
    canvas_showall[0].configure(scrollregion=canvas_showall[0].bbox("all"), width=800, height=500)


def showall_canvas_func1(event):
    canvas_showall[1].configure(scrollregion=canvas_showall[1].bbox("all"), width=800, height=500)


def showall_canvas_func2(event):
    canvas_showall[2].configure(scrollregion=canvas_showall[2].bbox("all"), width=800, height=500)


def editone_canvas_func(event):
    canvas_editone.configure(scrollregion=canvas_editone.bbox("all"), width=650, height=500)


def canvas_report_func(event):
    canvas_report.configure(scrollregion=canvas_report.bbox("all"), width=400, height=500)


def save_database():
    writer = pd.ExcelWriter(Database_Name, engine='xlsxwriter')
    data.to_excel(writer, sheet_name='Sheet1')
    writer.save()

    writer = pd.ExcelWriter(phars_Db_Name, engine='xlsxwriter')
    phars.to_excel(writer, sheet_name='Sheet1')
    writer.save()

    colfile = open("Database/column.txt", "w")
    colfile.write(str(colchar))
    colfile.close()

    counterfile = open("Database/counter.txt", "w")
    counterfile.write(str(fatora_counter))
    counterfile.close()


def on_closing2():
    save_database()
    tkrmsg.showinfo("Bye", "Thanks for using my Program\nDesigned By: Khaled Mohamed\nPhone: 01120982549\nE-mail: khaled.atya153@gmail.com")
    root.destroy()


def on_closing():
    if tkrmsg.askquestion("إغلاق", "سيتم فقدان الفاتورة بدون تسجيل\nهل انت متاكد ؟") == 'yes':
        fatora.destroy()
        rightframe_fatora.destroy()


def get_index(nname):
    for i in range(Database_Size):
        if data['Name'][i] == nname:
            return i

    return Database_Size


def find_in_phars(name):
    for i in range(phars_Size):
        if name == phars['Name'][i]:
            return True
    return False


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


def converttofloat(stri):
    stri = '0' + stri
    if '.' in stri:
        stri += '0'
    return float(stri)


def showall(event):
    global canvas_showall
    root.destroy()

    coun = (Database_Size//3)+1
    showall_window = [tkr.Tk() for i in range(3)]
    canvas_showall = [tkr.Canvas(showall_window[i]) for i in range(3)]
    in_canvas_showall = [tkr.Frame(canvas_showall[i]) for i in range(3)]
    myscrollbar = [tkr.Scrollbar(showall_window[i], orient="vertical", command=canvas_showall[i].yview) for i in range(3)]

    for i in range(3):
        showall_window[i].title('Page ' + str(i+1))
        showall_window[i].geometry('%dx%d+%d+%d' % (800, 500, 50*(3-i), 50*(3-i)))
        canvas_showall[i].configure(yscrollcommand=myscrollbar[i].set)
        myscrollbar[i].pack(side="right", fill="y")
        canvas_showall[i].pack(side="left")
        canvas_showall[i].create_window((0, 0), window=in_canvas_showall[i], anchor='nw')

        cpr_label = tkr.Label(in_canvas_showall[i], text='  سعر المستهلك  ', font=('Arial', '25'))
        name_label = tkr.Label(in_canvas_showall[i], text='  أسم الصنف  ', font=('Arial', '25'))
        bupr_label = tkr.Label(in_canvas_showall[i], text='  سعر الشراء ', font=('Arial', '25'))

        bupr_label.grid(row=0, column=0)
        name_label.grid(row=0, column=1)
        cpr_label.grid(row=0, column=2)

        for j in range(i*coun, min(Database_Size, (i+1)*coun)):
            cpr_label = tkr.Label(in_canvas_showall[i], text=("%.2f" % data['CPrice'][j]), font=('Arial', '15'))
            name_label = tkr.Label(in_canvas_showall[i], text=data['Name'][j], font=('Arial', '15'))
            bupr_label = tkr.Label(in_canvas_showall[i], text=("%.2f" % data['BuPrice'][j]), font=('Arial', '15'))

            bupr_label.grid(row=(j % coun)+1, column=0)
            name_label.grid(row=(j % coun)+1, column=1)
            cpr_label.grid(row=(j % coun)+1, column=2)
    in_canvas_showall[0].bind("<Configure>", showall_canvas_func0)
    in_canvas_showall[1].bind("<Configure>", showall_canvas_func1)
    in_canvas_showall[2].bind("<Configure>", showall_canvas_func2)
    showall_window[0].mainloop()
    main()


def addone(event):
    root.destroy()
    global cpr_entry_add, name_entry_add, bupr_entry_add
    addone_window = tkr.Tk()

    cpr_label = tkr.Label(addone_window, text='سعر المستهلك', font=('Arial', '15'))
    name_label = tkr.Label(addone_window, text='أسم الصنف', font=('Arial', '15'))
    bupr_label = tkr.Label(addone_window, text='سعر الشراء', font=('Arial', '15'))

    cpr_entry_add = tkr.Entry(addone_window, font=('Arial', '20'), width=5)
    name_entry_add = tkr.Entry(addone_window, font=('Arial', '20'), width=30)
    bupr_entry_add = tkr.Entry(addone_window, font=('Arial', '20'), width=5)

    add_button = tkr.Button(addone_window, text='إضافة', font=('Arial', '20'), width=10,
                            command=lambda m=0: add_button_pressed(m))

    bupr_label.grid(row=0)
    name_label.grid(row=0, column=1)
    cpr_label.grid(row=0, column=2)

    bupr_entry_add.grid(row=1, column=0)
    name_entry_add.grid(row=1, column=1)
    cpr_entry_add.grid(row=1, column=2)

    add_button.grid(row=2, columnspan=3)

    addone_window.mainloop()
    main()


def add_button_pressed(whocalled):
    if not cpr_entry_add.get() or not name_entry_add.get() or not bupr_entry_add.get():
        tkrmsg.showerror('خطأ', 'لا يمكن ترك احد الخانات فارغة')
        return
    if get_index(name_entry_add.get()) < Database_Size:
        tkrmsg.showerror('خطأ', 'هذا الأسم موجود من قبل')
        return
    if notnumber(cpr_entry_add.get()):
        tkrmsg.showerror('خطأ', 'خطأ فى سعر المستهلك')
        return
    if notnumber(bupr_entry_add.get()):
        tkrmsg.showerror('خطأ', 'خطأ فى سعر الشراء')
        return

    tt1 = converttofloat(cpr_entry_add.get())
    tt2 = converttofloat(bupr_entry_add.get())
    add_product(name_entry_add.get(), tt1, tt2)
    save_database()
    tkrmsg.showinfo('تم', 'تم اضافة الصنف بنجاح')
    if whocalled:
        update_edit('', '', '', 1)


def add_product(name, cpr, bupr):
    global Database_Size
    listofzeros = [0] * colchar
    listofzeros[0] = name
    listofzeros[1] = cpr
    listofzeros[2] = bupr
    data.loc[Database_Size] = listofzeros
    Database_Size += 1


def editone(event):
    global in_canvas_editone, canvas_editone
    root.destroy()
    editone_window = tkr.Tk()

    topframe = tkr.Frame(editone_window)
    topframe.pack()

    canvas_editone = tkr.Canvas(editone_window, height=2)
    in_canvas_editone = tkr.Frame(canvas_editone, height=2)

    myscrollbar = tkr.Scrollbar(editone_window, orient="vertical", command=canvas_editone.yview)
    canvas_editone.configure(yscrollcommand=myscrollbar.set)
    myscrollbar.pack(side="right", fill="y")
    canvas_editone.pack(side="left")
    canvas_editone.create_window((0, 0), window=in_canvas_editone, anchor='nw')
    in_canvas_editone.bind("<Configure>", editone_canvas_func)

    cur_name = tkr.StringVar()
    cur_cpr = tkr.StringVar()
    cur_bupr = tkr.StringVar()

    cur_name.trace("w", lambda name, index, mode: update_edit(cur_name.get(), cur_cpr.get(), cur_bupr.get(), 0))
    cur_cpr.trace("w", lambda name, index, mode: update_edit(cur_name.get(), cur_cpr.get(), cur_bupr.get(), 0))
    cur_bupr.trace("w", lambda name, index, mode: update_edit(cur_name.get(), cur_cpr.get(), cur_bupr.get(), 0))

    cpr_label = tkr.Label(topframe, text='  سعر المستهلك  ', font=('Arial', '15'))
    name_label = tkr.Label(topframe, text='  أسم الصنف  ', font=('Arial', '15'))
    bupr_label = tkr.Label(topframe, text='سعر الشراء', font=('Arial', '15'))
    emp_label = tkr.Label(topframe, text='', font=('Arial', '15'), width=10)
    bupr_label.grid(row=0, column=0)
    name_label.grid(row=0, column=1)
    cpr_label.grid(row=0, column=2)
    emp_label.grid(row=0, column=3)

    cpr_entry_edit = tkr.Entry(topframe, font=('Arial', '20'), width=5, textvariable=cur_cpr)
    name_entry_edit = tkr.Entry(topframe, font=('Arial', '20'), width=23, textvariable=cur_name)
    bupr_entry_edit = tkr.Entry(topframe, font=('Arial', '20'), width=5, textvariable=cur_bupr)

    bupr_entry_edit.grid(row=1, column=0)
    name_entry_edit.grid(row=1, column=1)
    cpr_entry_edit.grid(row=1, column=2)

    editone_window.mainloop()
    main()


def update_edit(name, cpr, bupr, whocalled):
    if whocalled:
        fr = in_canvas_downframe_fatora
        canvas_downframe_fatora.yview_moveto('0.0')
    else:
        fr = in_canvas_editone
        canvas_editone.yview_moveto('0.0')

    for thing in fr.grid_slaves():
        thing.grid_forget()

    if notnumber(cpr) or notnumber(bupr) or cpr == '.' or bupr == '.':
        return

    if (not cpr) and (not bupr) and (len(name) < 3):
        return

    fl1 = 0
    fl2 = 0

    if not cpr:
        fl1 = 1
    else:
        cpr = converttofloat(cpr)

    if not bupr:
        fl2 = 1
    else:
        bupr = converttofloat(bupr)

    for i in range(Database_Size):
        t = data.loc[i]
        if (name.lower() in t["Name"].lower()) and (fl1 or (cpr == float(t["CPrice"]))) and (fl2 or (bupr == float(t["BuPrice"]))):
            cpr_label = tkr.Label(fr, text=("%.2f" % t['CPrice']) + '  ', font=('Arial', '15'), width=7)
            bupr_label = tkr.Label(fr, text=("%.2f" % t['BuPrice']), font=('Arial', '15'), width=5)

            if whocalled:
                name_label = tkr.Button(fr, text=t['Name'], font=('Arial', '15'), width=35, padx=10,
                                        command=lambda m=t["Name"], mm=t["CPrice"], mmm=i: addinfatora_pressed(m, mm, mmm))
                edit_button = tkr.Button(fr, text='تعديل', font=('Arial', '15'),
                                         command=lambda m=i, mm=1: edit_product(m, mm), width=3, padx=10)
                phpr_label = tkr.Label(fr, text=("    %.2f" % t[phname_entry_get]), font=('Arial', '15'), width=10)
                phpr_label.grid(row=i, column=6)
            else:
                name_label = tkr.Label(fr, text=t['Name'], font=('Arial', '15'), width=35)
                remove_button = tkr.Button(fr, text='حذف', font=('Arial', '20'),
                                           command=lambda m=i: delete_product_conf(m), width=3)
                edit_button = tkr.Button(fr, text='تعديل', font=('Arial', '20'),
                                         command=lambda m=i, mm=0: edit_product(m, mm), width=3)
                remove_button.grid(row=i, column=4)

            bupr_label.grid(row=i, column=0)
            name_label.grid(row=i, column=1)
            cpr_label.grid(row=i, column=2)
            edit_button.grid(row=i, column=3)


def filter_phar(name):
    canvas_report.yview_moveto('0.0')
    for thing in in_canvas_report.grid_slaves():
        thing.grid_forget()

    if len(name) < 3:
        return

    for i in range(phars_Size):
        if name.lower() in phars["Name"][i].lower():
            empty_label = tkr.Label(in_canvas_report, text=' ', width=10)
            phar_chose_button = tkr.Button(in_canvas_report, text=phars["Name"][i], font=('Arial', '15'),
                                           command=lambda m=phars["Name"][i]: phar_chosen(m), width=20)
            empty_label.grid(row=i, column=0)
            phar_chose_button.grid(row=i, column=1)


def phar_chosen(name):
    pharm_name.set(name)


def delete_product_conf(found):
    answer = tkrmsg.askquestion("تأكيد", "حذف " + data['Name'][found] + '\n هل انت متأكد ؟')
    if answer == 'yes':
        delete_product(found)
        save_database()
        update_edit('', '', '', 0)


def delete_product(found):
    global Database_Size
    Database_Size -= 1
    data.loc[found] = deepcopy(data.loc[Database_Size])
    data.drop(data.index[Database_Size], inplace=True)


def edit_product(found, whocalled):
    global cpr_entry_edit2, name_entry_edit2, bupr_entry_edit2, edit_window

    edit_window = tkr.Tk()

    cpr_label_edit2 = tkr.Label(edit_window, text='سعر المستهلك', font=('Arial', '15'))
    name_label_edit2 = tkr.Label(edit_window, text='أسم الصنف', font=('Arial', '15'))
    bupr_label_edit2 = tkr.Label(edit_window, text='سعر الشراء', font=('Arial', '15'))

    cpr_entry_edit2 = tkr.Entry(edit_window, font=('Arial', '20'), width=5)
    name_entry_edit2 = tkr.Entry(edit_window, font=('Arial', '20'), width=30)
    bupr_entry_edit2 = tkr.Entry(edit_window, font=('Arial', '20'), width=5)

    cpr_entry_edit2.insert(0, data['CPrice'][found])
    name_entry_edit2.insert(0, data['Name'][found])
    bupr_entry_edit2.insert(0, data['BuPrice'][found])

    edit_button2 = tkr.Button(edit_window, text='تعديل', font=('Arial', '20'),
                              width=10, command=lambda m=found, wh=whocalled: edit_buttpon_clicked(m, wh))

    bupr_label_edit2.grid(row=0)
    name_label_edit2.grid(row=0, column=1)
    cpr_label_edit2.grid(row=0, column=2)

    bupr_entry_edit2.grid(row=1, column=0)
    name_entry_edit2.grid(row=1, column=1)
    cpr_entry_edit2.grid(row=1, column=2)

    edit_button2.grid(row=2, columnspan=3)

    edit_window.mainloop()


def edit_buttpon_clicked(idx, whocalled):
    if not cpr_entry_edit2.get() or not name_entry_edit2.get() or not bupr_entry_edit2.get():
        tkrmsg.showerror('خطأ', 'لا يمكن ترك احد الخانات فارغة')
        return
    if (get_index(name_entry_edit2.get()) < Database_Size) and (get_index(name_entry_edit2.get()) != idx):
        tkrmsg.showerror('خطأ', 'هذا الأسم موجود من قبل')
        return
    if notnumber(cpr_entry_edit2.get()):
        tkrmsg.showerror('خطأ', 'خطأ فى سعر المستهلك')
        return
    if notnumber(bupr_entry_edit2.get()):
        tkrmsg.showerror('خطأ', 'خطأ فى سعر الصيدلي')
        return

    tt1 = converttofloat(cpr_entry_edit2.get())
    tt2 = converttofloat(bupr_entry_edit2.get())

    data.at[idx, "Name"] = name_entry_edit2.get()
    data.at[idx, "CPrice"] = float(tt1)
    data.at[idx, "BuPrice"] = float(tt2)

    save_database()
    tkrmsg.showinfo('تم', 'تم التعديل')
    edit_window.destroy()
    update_edit('', '', '', whocalled)


def report(event):
    global date_ph, phname_entry, canvas_report, in_canvas_report, pharm_name
    root.destroy()
    date_ph = tkr.Tk()

    topframe = tkr.Frame(date_ph)
    topframe.pack()

    canvas_report = tkr.Canvas(date_ph, height=2)
    in_canvas_report = tkr.Frame(canvas_report, height=2)

    myscrollbar = tkr.Scrollbar(date_ph, orient="vertical", command=canvas_report.yview)
    canvas_report.configure(yscrollcommand=myscrollbar.set)
    myscrollbar.pack(side="right", fill="y")
    canvas_report.pack(side="left")
    canvas_report.create_window((0, 0), window=in_canvas_report, anchor='nw')
    in_canvas_report.bind("<Configure>", canvas_report_func)
    pharm_name = tkr.StringVar()
    pharm_name.trace("w", lambda name, index, mode: filter_phar(pharm_name.get()))

    phname_label = tkr.Label(topframe, text=' أسم الصيدلية ', font=('Arial', '15'))
    phname_entry = tkr.Entry(topframe, font=('Arial', '15'), width=20, textvariable=pharm_name)
    start_button = tkr.Button(topframe, text='بدء', font=('Arial', '15'), width=7,  command=start_buttpon_clicked)

    start_button.grid(row=0, column=0)
    phname_entry.grid(row=0, column=1)
    phname_label.grid(row=0, column=2)

    date_ph.mainloop()
    main()


def start_buttpon_clicked():
    global downframe_fatora, quantity_entry, phpr_entry, in_canvas_rightframe_fatora, total_label, phname_entry_get,\
        df_fat, sz_fat, canvas_rightframe_fatora, bupr_entry_add, name_entry_add, cpr_entry_add, fatora_total, \
        in_canvas_downframe_fatora, count_label, fatora, rightframe_fatora, phars_Size, canvas_downframe_fatora, colchar

    fatora_total = 0
    sz_fat = 0
    df_fat = pd.DataFrame(columns=['Name', 'Quantity', 'PhPrice', 'CPrice'])

    phname_entry_get = phname_entry.get()
    if not phname_entry_get:
        phname_entry_get = '-----------------'
        answer = tkrmsg.askquestion('تحذير', "لم يتم ادخال اسم الصيدلية!\n متأكد انك تريد المتابعة ؟")
        if answer == 'no':
            return

    if (not find_in_phars(phname_entry_get)) and (phname_entry_get != '-----------------'):
        phars.loc[phars_Size] = [phname_entry_get]
        save_database()
        phars_Size += 1

    if phname_entry_get not in data.columns:
        listofzeros = [0] * Database_Size
        data[phname_entry_get] = listofzeros
        colchar += 1
        save_database()

    phname_entry_get = phname_entry_get.replace('/', '.')
    phname_entry_get = phname_entry_get.replace('\\', '.')

    date_ph.destroy()
    fatora = tkr.Tk()
    rightframe_fatora = tkr.Tk()
    fatora.title(phname_entry_get)
    rightframe_fatora.title(phname_entry_get)

    ws = rightframe_fatora.winfo_screenwidth()
    hs = rightframe_fatora.winfo_screenheight()

    fatora.geometry('%dx%d+%d+%d' % (760, 625, 0, 0))
    rightframe_fatora.geometry('%dx%d+%d+%d' % (700, 500, ws - 715, hs - 600))

    fatora.protocol("WM_DELETE_WINDOW", on_closing)
    rightframe_fatora.protocol("WM_DELETE_WINDOW", on_closing)
    topframe_fatora = tkr.Frame(fatora)
    downframe_fatora = tkr.Frame(fatora)

    topframe_fatora.grid(row=0, column=0)
    downframe_fatora.grid(row=1, column=0)

    canvas_downframe_fatora = tkr.Canvas(downframe_fatora)
    in_canvas_downframe_fatora = tkr.Frame(canvas_downframe_fatora)

    myscrollbar = tkr.Scrollbar(downframe_fatora, orient="vertical", command=canvas_downframe_fatora.yview)
    canvas_downframe_fatora.configure(yscrollcommand=myscrollbar.set)
    myscrollbar.pack(side="right", fill="y")
    canvas_downframe_fatora.pack(side="left")
    canvas_downframe_fatora.create_window((0, 0), window=in_canvas_downframe_fatora, anchor='nw')
    in_canvas_downframe_fatora.bind("<Configure>", canvasfunc2)

    canvas_rightframe_fatora = tkr.Canvas(rightframe_fatora)
    in_canvas_rightframe_fatora = tkr.Frame(canvas_rightframe_fatora)

    myscrollbar = tkr.Scrollbar(rightframe_fatora, orient="vertical", command=canvas_rightframe_fatora.yview)
    canvas_rightframe_fatora.configure(yscrollcommand=myscrollbar.set)
    myscrollbar.pack(side="right", fill="y")
    canvas_rightframe_fatora.pack(side="left")
    canvas_rightframe_fatora.create_window((0, 0), window=in_canvas_rightframe_fatora, anchor='nw')
    in_canvas_rightframe_fatora.bind("<Configure>", canvasfunc)

    cur_name = tkr.StringVar()
    cur_cpr = tkr.StringVar()
    cur_bupr = tkr.StringVar()

    cur_name.trace("w", lambda name, index, mode: update_edit(cur_name.get(), cur_cpr.get(), cur_bupr.get(), 1))
    cur_cpr.trace("w", lambda name, index, mode: update_edit(cur_name.get(), cur_cpr.get(), cur_bupr.get(), 1))
    cur_bupr.trace("w", lambda name, index, mode: update_edit(cur_name.get(), cur_cpr.get(), cur_bupr.get(), 1))

    cpr_label = tkr.Label(topframe_fatora, text='  سعر المستهلك  ', font=('Arial', '15'))
    name_label = tkr.Label(topframe_fatora, text='  أسم الصنف  ', font=('Arial', '15'))
    bupr_label = tkr.Label(topframe_fatora, text='سعر الشراء', font=('Arial', '15'))
    quantity_label = tkr.Label(topframe_fatora, text='  الكمية  ', font=('Arial', '15'))
    phpr_label = tkr.Label(topframe_fatora, text='  سعر الصيدلي  ', font=('Arial', '15'))

    bupr_entry_fatora = tkr.Entry(topframe_fatora, font=('Arial', '20'), width=5, textvariable=cur_bupr)
    name_entry_fatora = tkr.Entry(topframe_fatora, font=('Arial', '20'), width=24, textvariable=cur_name)
    cpr_entry_fatora = tkr.Entry(topframe_fatora, font=('Arial', '20'), width=5, textvariable=cur_cpr)
    quantity_entry = tkr.Entry(topframe_fatora, font=('Arial', '20'), width=5)
    phpr_entry = tkr.Entry(topframe_fatora, font=('Arial', '20'), width=5)

    bupr_entry_add = tkr.Entry(topframe_fatora, font=('Arial', '20'), width=5)
    name_entry_add = tkr.Entry(topframe_fatora, font=('Arial', '20'), width=24)
    cpr_entry_add = tkr.Entry(topframe_fatora, font=('Arial', '20'), width=5)
    add_button_inside_fatora = tkr.Button(topframe_fatora, text='إضافة صنف جديد',
                                          font=('Arial', '15'), command=lambda m=1: add_button_pressed(m))

    bupr_entry_add.grid(row=0, column=0)
    name_entry_add.grid(row=0, column=1, columnspan=3)
    cpr_entry_add.grid(row=0, column=4)
    add_button_inside_fatora.grid(row=0, column=5, columnspan=2)

    bupr_label.grid(row=1, column=0)
    name_label.grid(row=1, column=1, columnspan=3)
    cpr_label.grid(row=1, column=4)
    quantity_label.grid(row=1, column=5)
    phpr_label.grid(row=1, column=6)

    bupr_entry_fatora.grid(row=2, column=0)
    name_entry_fatora.grid(row=2, column=1, columnspan=3)
    cpr_entry_fatora.grid(row=2, column=4)
    quantity_entry.grid(row=2, column=5)
    phpr_entry.grid(row=2, column=6)

    count_label = tkr.Label(in_canvas_rightframe_fatora, text='0', font=('Arial', '15'))
    countname_label = tkr.Label(in_canvas_rightframe_fatora, text='  عدد الأصناف  ', font=('Arial', '15'))
    total_label = tkr.Label(in_canvas_rightframe_fatora, text='0.00', font=('Arial', '15'))
    totalname_label = tkr.Label(in_canvas_rightframe_fatora, text='  الإجمالى  ', font=('Arial', '15'))
    save_fatora_button = tkr.Button(in_canvas_rightframe_fatora, text='حفظ', font=('Arial', '17'), width=5,
                                    padx=5, command=save_pressed)

    total_label.grid(row=0, column=0)
    totalname_label.grid(row=0, column=1)
    count_label.grid(row=0, column=2)
    countname_label.grid(row=0, column=3)
    save_fatora_button.grid(row=0, column=4)

    name1_label = tkr.Label(in_canvas_rightframe_fatora, text='  أسم الصنف  ', font=('Arial', '15'), width=28)
    phpr1_label = tkr.Label(in_canvas_rightframe_fatora, text='  سعر القطعة  ', font=('Arial', '15'))
    quantity1_label = tkr.Label(in_canvas_rightframe_fatora, text='  الكمية  ', font=('Arial', '15'))
    phprtotal1_label = tkr.Label(in_canvas_rightframe_fatora, text='  السعر الكلي  ', font=('Arial', '15'))

    phprtotal1_label.grid(row=1, column=0)
    phpr1_label.grid(row=1, column=1)
    quantity1_label.grid(row=1, column=2)
    name1_label.grid(row=1, column=3)
    rightframe_fatora.mainloop()
    fatora.mainloop()


def addinfatora_pressed(foundname, foundcpr, idxx):
    global sz_fat, fatora_total
    quan = quantity_entry.get()
    phprprice = '0' + phpr_entry.get()
    if '.' in phprprice:
        phprprice += '0'

    if (notnumber(quan)) or ('.' in quan) or (not quan) or (not int(quan)):
        tkrmsg.showerror("خطأ", 'خطأ فى الكمية .. لم يتم اضافة الصنف إلى الفاتورة')
        return

    if (notnumber(phprprice)) or (not phprprice) or (not float(phprprice)):
        tkrmsg.showerror("خطأ", 'خطأ فى سعر الصيدلي .. لم يتم اضافة الصنف إلى الفاتورة')
        return

    quan = int(quan)
    phprprice = float(phprprice)

    data.at[idxx, phname_entry_get] = phprprice

    df_fat.loc[sz_fat] = [foundname, quan, phprprice, foundcpr]

    add = phprprice*quan
    name2_label = tkr.Label(in_canvas_rightframe_fatora, text=foundname,
                            font=('Arial', '15'), width=28)
    phpr2_label = tkr.Label(in_canvas_rightframe_fatora, text=("%.2f" % phprprice), font=('Arial', '15'))
    quantity2_label = tkr.Label(in_canvas_rightframe_fatora, text=quan, font=('Arial', '15'))
    phprtotal2_label = tkr.Label(in_canvas_rightframe_fatora, text=("%.2f" % add), font=('Arial', '15'))
    removefatora_button = tkr.Button(in_canvas_rightframe_fatora, text='حذف',
                                     font=('Arial', '15'), command=lambda m=sz_fat: removefatora_pressed(m))

    fatora_total += add

    phprtotal2_label.grid(row=2+sz_fat, column=0)
    phpr2_label.grid(row=2+sz_fat, column=1)
    quantity2_label.grid(row=2+sz_fat, column=2)
    name2_label.grid(row=2+sz_fat, column=3)
    removefatora_button.grid(row=2+sz_fat, column=4)

    sz_fat += 1
    total_label['text'] = ("%.2f" % fatora_total)
    count_label['text'] = sz_fat


def removefatora_pressed(idx):
    global sz_fat, fatora_total
    answer = tkrmsg.askquestion("تأكيد", "حذف " + df_fat['Name'][idx] + '\n هل انت متأكد ؟')
    if answer == 'no':
        return

    add = df_fat['PhPrice'][idx]*df_fat['Quantity'][idx]
    fatora_total -= add

    for i in range(idx+1, sz_fat):
        df_fat.loc[i-1] = df_fat.loc[i]
        add = df_fat['PhPrice'][i] * df_fat['Quantity'][i]

        name2_label = tkr.Label(in_canvas_rightframe_fatora, text=df_fat['Name'][i],
                                font=('Arial', '15'), width=28)
        phpr2_label = tkr.Label(in_canvas_rightframe_fatora, text=("%.2f" % df_fat['PhPrice'][i]), font=('Arial', '15'))
        quantity2_label = tkr.Label(in_canvas_rightframe_fatora, text=df_fat['Quantity'][i], font=('Arial', '15'))
        phprtotal2_label = tkr.Label(in_canvas_rightframe_fatora, text=("%.2f" % add), font=('Arial', '15'))
        removefatora_button = tkr.Button(in_canvas_rightframe_fatora, text='حذف',
                                         font=('Arial', '15'), command=lambda m=i-1: removefatora_pressed(m))

        phprtotal2_label.grid(row=2+i-1, column=0)
        phpr2_label.grid(row=2+i-1, column=1)
        quantity2_label.grid(row=2+i-1, column=2)
        name2_label.grid(row=2+i-1, column=3)
        removefatora_button.grid(row=2+i-1, column=4)

    sz_fat -= 1
    df_fat.drop(df_fat.index[sz_fat], inplace=True)

    for thing in in_canvas_rightframe_fatora.grid_slaves(row=2+sz_fat):
        #  if int(thing.grid_info()["row"]) == 2+sz_fat:
        thing.grid_forget()

    total_label['text'] = ("%.2f" % fatora_total)
    count_label['text'] = sz_fat


"""
def show_df_fat():
    global fatora_total
    total = 0

    for thing in in_canvas_rightframe_fatora.grid_slaves():
        if int(thing.grid_info()["row"]) > 1:
            thing.grid_forget()

    for i in range(sz_fat):
        new_pr = float(df_fat['PhPrice'][i])
        qua = int(df_fat['Quantity'][i])
        add = new_pr*qua
        name2_label = tkr.Label(in_canvas_rightframe_fatora, text=df_fat['Name'][i],
                                font=('Arial', '15'), width=28)
        phpr2_label = tkr.Label(in_canvas_rightframe_fatora, text=("%.2f" % new_pr), font=('Arial', '15'))
        quantity2_label = tkr.Label(in_canvas_rightframe_fatora, text=qua, font=('Arial', '15'))
        phprtotal2_label = tkr.Label(in_canvas_rightframe_fatora, text=("%.2f" % add), font=('Arial', '15'))
        removefatora_button = tkr.Button(in_canvas_rightframe_fatora, text='حذف',
                                         font=('Arial', '15'), command=lambda m=i: removefatora_pressed(m))
        total += add

        phprtotal2_label.grid(row=2+i, column=0)
        phpr2_label.grid(row=2+i, column=1)
        quantity2_label.grid(row=2+i, column=2)
        name2_label.grid(row=2+i, column=3)
        removefatora_button.grid(row=2+i, column=4)

    fatora_total = total
    total_label['text'] = ("%.2f" % fatora_total)
    count_label['text'] = len(df_fat)
"""


def save_pressed():
    global dire, fatora_counter
    if not sz_fat:
        tkrmsg.showerror('خطأ', 'لم يتم إدخال اى صنف\nلا يمكن طباعة الفاتورة فارغة')
        return

    ans = tkrmsg.askquestion('تأكيد', 'هل انت متأكد من انك تريد حفظ الفاتورة للطباعة ؟')
    if ans == 'no':
        return

    fatora_counter += 1
    save_database()

    dire = 'Reports\\' + phname_entry_get + '\\' + str(date.today()) + '\\' + str(fatora_counter) + '\\'
    if not os.path.exists(dire):
        os.makedirs(dire)

    for page in range(1, (sz_fat + 49)//25):
        # total = 0
        template = openpyxl.load_workbook('Database/template.xlsx')
        sheet = template['Sheet1']
        sheet['A3'] = date.today()
        sheet['E3'] = 'بيان تسليم رقم ' + str(fatora_counter) + ' صفحة ' + str(page)
        sheet['E5'] = 'البضاعة مسلمة إلي : ' + phname_entry_get

        for i in range((page - 1)*25, min(page * 25, sz_fat)):
            qua = int(df_fat['Quantity'][i])
            fatpr = float(df_fat['PhPrice'][i])

            name = df_fat['Name'][i]
            cupr = df_fat['CPrice'][i]

            # sheet['A' + str(9 + (i % 25))] = ("%.2f" % add)
            sheet['B' + str(9 + (i % 25))] = ("%.2f" % fatpr)
            if qua % 12 != 0:
                sheet['C' + str(9 + (i % 25))] = 0
                sheet['D' + str(9 + (i % 25))] = qua
            else:
                sheet['C' + str(9 + (i % 25))] = qua//12
                sheet['D' + str(9 + (i % 25))] = 0
            sheet['E' + str(9 + (i % 25))] = name
            sheet['F' + str(9 + (i % 25))] = ("%.2f" % cupr)
            sheet['G' + str(9 + (i % 25))] = (i % 25)+1

        # sheet['A34'] = ("%.2f" % total)
        template.save(dire + str(page) + '.xlsx')
        subprocess.Popen(r'explorer /open,"' + dire + str(page) + '.xlsx')

    save_records()

    fatora.destroy()
    rightframe_fatora.destroy()


def save_records():
    os.makedirs(dire + 'DataFrame')
    writer = pd.ExcelWriter(dire + 'DataFrame\\fatora.xlsx', engine='xlsxwriter')
    df_fat.to_excel(writer, sheet_name='Sheet1')
    writer.save()


start()
