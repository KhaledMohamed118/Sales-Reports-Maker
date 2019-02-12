import subprocess
import os
import openpyxl
import tkinter as tkr
import tkinter.messagebox as tkrmsg
import pandas as pd
import numpy
from copy import deepcopy
from datetime import date
from openpyxl.styles.borders import Border, Side


def main():
    global data, root, Database_Size, Database_Name
    Database_Name = "Database/Book2.xlsx"
    data = pd.read_excel(Database_Name, usecols="B:D")
    Database_Size = len(data)
    save_database()
    root = tkr.Tk()
    root.title('Mr/Abd Elrahman')

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
    canvas_rightframe_fatora.configure(scrollregion=canvas_rightframe_fatora.bbox("all"), width=525, height=500)


def canvasfunc2(event):
    canvas_downframe_fatora.configure(scrollregion=canvas_downframe_fatora.bbox("all"), width=700, height=500)


def showall_canvas_func(event):
    canvas_showall.configure(scrollregion=canvas_showall.bbox("all"), width=700, height=500)


def editone_canvas_func(event):
    canvas_editone.configure(scrollregion=canvas_editone.bbox("all"), width=800, height=500)


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


def search(name, cpr, bupr):
    if notnumber(cpr):
        cpr = 0
    else:
        cpr = '0' + cpr
        if '.' in cpr:
            cpr += '0'
        cpr = float(cpr)

    if notnumber(bupr):
        bupr = 0
    else:
        bupr = '0' + bupr
        if '.' in bupr:
            bupr += '0'
        bupr = float(bupr)

    sz = 0
    df = pd.DataFrame(columns=['Name', 'CPrice', 'BuPrice'])

    for i in range(Database_Size):
        t = data.loc[i]
        if (name.lower() in t["Name"].lower()) and ((not cpr) or (cpr == float(t["CPrice"]))) and ((not bupr) or (bupr == float(t["BuPrice"]))):
            df.loc[sz] = data.loc[i]
            sz += 1

    return df


def showall(event):
    global canvas_showall
    root.destroy()
    showall_window = tkr.Tk()

    canvas_showall = tkr.Canvas(showall_window)
    in_canvas_showall = tkr.Frame(canvas_showall)

    myscrollbar = tkr.Scrollbar(showall_window, orient="vertical", command=canvas_showall.yview)
    canvas_showall.configure(yscrollcommand=myscrollbar.set)
    myscrollbar.pack(side="right", fill="y")
    canvas_showall.pack(side="left")
    canvas_showall.create_window((0, 0), window=in_canvas_showall, anchor='nw')
    in_canvas_showall.bind("<Configure>", showall_canvas_func)

    cpr_label = tkr.Label(in_canvas_showall, text='  سعر المستهلك  ', font=('Arial', '25'))
    name_label = tkr.Label(in_canvas_showall, text='  أسم الصنف  ', font=('Arial', '25'))
    bupr_label = tkr.Label(in_canvas_showall, text='  سعر الشراء ', font=('Arial', '25'))

    bupr_label.grid(row=0, column=0)
    name_label.grid(row=0, column=1)
    cpr_label.grid(row=0, column=2)

    for i in range(Database_Size):
        cpr_label = tkr.Label(in_canvas_showall, text=("%.2f" % data['CPrice'][i]), font=('Arial', '15'))
        name_label = tkr.Label(in_canvas_showall, text=data['Name'][i], font=('Arial', '15'))
        bupr_label = tkr.Label(in_canvas_showall, text=("%.2f" % data['BuPrice'][i]), font=('Arial', '15'))

        bupr_label.grid(row=i+1, column=0)
        name_label.grid(row=i+1, column=1)
        cpr_label.grid(row=i+1, column=2)
    showall_window.mainloop()
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

    add_button = tkr.Button(addone_window, text='إضافة', font=('Arial', '20'), width=10, command=add_button_pressed)
    # add_button.flash()
    # add_button.bind("<Button-1>", add_button_pressed)

    bupr_label.grid(row=0)
    name_label.grid(row=0, column=1)
    cpr_label.grid(row=0, column=2)

    bupr_entry_add.grid(row=1, column=0)
    name_entry_add.grid(row=1, column=1)
    cpr_entry_add.grid(row=1, column=2)

    add_button.grid(row=2, columnspan=3)

    addone_window.mainloop()
    main()


def add_button_pressed():
    if not cpr_entry_add.get() or not name_entry_add.get() or not bupr_entry_add.get():
        tkrmsg.showerror('خطأ', 'لا يمكن ترك احد الخانات فارغة')

    elif get_index(name_entry_add.get()) < Database_Size:
        tkrmsg.showerror('خطأ', 'هذا الأسم موجود من قبل')

    else:
        if notnumber(cpr_entry_add.get()):
            tkrmsg.showerror('خطأ', 'خطأ فى سعر المستهلك')
        elif notnumber(bupr_entry_add.get()):
            tkrmsg.showerror('خطأ', 'خطأ فى سعر الشراء')
        else:
            tt1 = '0' + cpr_entry_add.get()
            tt2 = '0' + bupr_entry_add.get()
            if '.' in tt1:
                tt1 += '0'
            if '.' in tt2:
                tt2 += '0'
            add_product(name_entry_add.get(), float(tt1), float(tt2))
            tkrmsg.showinfo('تم', 'تم اضافة الصنف بنجاح')


def add_product(name, cpr, bupr):
    global Database_Size
    data.loc[Database_Size] = [name, cpr, bupr]
    """try:
    	data.loc[Database_Size] = [name, cpr, bupr]
    except:
    	print('err')
    	data.append([name, cpr, bupr])"""

    Database_Size += 1
    save_database()


def editone(event):
    global in_canvas_editone, canvas_editone
    root.destroy()
    editone_window = tkr.Tk()

    topframe = tkr.Frame(editone_window)
    topframe.pack()

    canvas_editone = tkr.Canvas(editone_window)
    in_canvas_editone = tkr.Frame(canvas_editone)

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

    bupr_label.grid(row=0, column=0)
    name_label.grid(row=0, column=1)
    cpr_label.grid(row=0, column=2)

    cpr_entry_edit = tkr.Entry(topframe, font=('Arial', '20'), width=5, textvariable=cur_cpr)
    name_entry_edit = tkr.Entry(topframe, font=('Arial', '20'), width=30, textvariable=cur_name)
    bupr_entry_edit = tkr.Entry(topframe, font=('Arial', '20'), width=5, textvariable=cur_bupr)

    bupr_entry_edit.grid(row=1, column=0)
    name_entry_edit.grid(row=1, column=1)
    cpr_entry_edit.grid(row=1, column=2)
    update_edit('', '', '', 0)
    editone_window.mainloop()
    main()


def update_edit(name, cpr, bupr, whocalled):
    if whocalled:
        fr = in_canvas_downframe_fatora
    else:
        fr = in_canvas_editone

    for thing in fr.grid_slaves():
        thing.grid_forget()

    df = search(name, cpr, bupr)
    for i in range(len(df)):
        cpr_label = tkr.Label(fr, text=("%.2f" % df['CPrice'][i]), font=('Arial', '15'), width=4)
        name_label = tkr.Label(fr, text=df['Name'][i], font=('Arial', '15'), width=40)
        bupr_label = tkr.Label(fr, text=("%.2f" % df['BuPrice'][i]), font=('Arial', '15'), width=4)

        bupr_label.grid(row=i, column=0)
        name_label.grid(row=i, column=1)
        cpr_label.grid(row=i, column=2)

        if whocalled:
            addinfatora_button = tkr.Button(fr, text='إضافة', font=('Arial', '20'), command=lambda m=df['Name'][i]: addinfatora_pressed(m), width=3)
            addinfatora_button.grid(row=i, column=4)
        else:
            edit_button = tkr.Button(fr, text='تعديل', font=('Arial', '20'), command=lambda m=df['Name'][i]: edit_product(m), width=3)
            remove_button = tkr.Button(fr, text='حذف', font=('Arial', '20'), command=lambda m=df['Name'][i]: delete_product(m), width=3)

            edit_button.grid(row=i, column=3)
            remove_button.grid(row=i, column=4)


def delete_product(name):
    answer = tkrmsg.askquestion("تأكيد", "حذف " + name + '\n هل انت متأكد ؟')
    if answer == 'no':
        return
    global Database_Size
    found = get_index(name)

    for i in range(found+1, Database_Size):
        data.loc[i-1] = data.loc[i]
    Database_Size -= 1
    data.drop(data.index[Database_Size], inplace=True)
    save_database()
    update_edit('', '', '', 0)


def edit_product(name):
    global cpr_entry_edit2, name_entry_edit2, bupr_entry_edit2, edit_window

    found = get_index(name)

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
                              width=10, command=lambda m=found: edit_buttpon_clicked(m))

    bupr_label_edit2.grid(row=0)
    name_label_edit2.grid(row=0, column=1)
    cpr_label_edit2.grid(row=0, column=2)

    bupr_entry_edit2.grid(row=1, column=0)
    name_entry_edit2.grid(row=1, column=1)
    cpr_entry_edit2.grid(row=1, column=2)

    edit_button2.grid(row=2, columnspan=3)

    edit_window.mainloop()


def edit_buttpon_clicked(idx):
    if not cpr_entry_edit2.get() or not name_entry_edit2.get() or not bupr_entry_edit2.get():
        tkrmsg.showerror('خطأ', 'لا يمكن ترك احد الخانات فارغة')

    elif (get_index(name_entry_edit2.get()) < Database_Size) and (get_index(name_entry_edit2.get()) != idx):
        tkrmsg.showerror('خطأ', 'هذا الأسم موجود من قبل')

    else:
        if notnumber(cpr_entry_edit2.get()):
            tkrmsg.showerror('خطأ', 'خطأ فى سعر المستهلك')
        elif notnumber(bupr_entry_edit2.get()):
            tkrmsg.showerror('خطأ', 'خطأ فى سعر الصيدلي')
        else:
            tt1 = '0' + cpr_entry_edit2.get()
            tt2 = '0' + bupr_entry_edit2.get()
            if '.' in tt1:
                tt1 += '0'
            if '.' in tt2:
                tt2 += '0'

            data['Name'][idx] = name_entry_edit2.get()
            data['CPrice'][idx] = float(tt1)
            data['BuPrice'][idx] = float(tt2)
            tkrmsg.showinfo('تم', 'تم التعديل')
            save_database()
            edit_window.destroy()
            update_edit('', '', '', 0)


def report(event):
    global date_ph, phname_entry
    root.destroy()
    date_ph = tkr.Tk()

    phname_label = tkr.Label(date_ph, text=' أسم الصيدلية ', font=('Arial', '15'))
    phname_entry = tkr.Entry(date_ph, font=('Arial', '20'), width=30)
    start_button = tkr.Button(date_ph, text='بدء', font=('Arial', '20'), width=10,  command=start_buttpon_clicked)

    phname_entry.grid(row=0, column=0)
    phname_label.grid(row=0, column=1)
    start_button.grid(row=1, columnspan=2)

    date_ph.mainloop()
    main()


def start_buttpon_clicked():
    global downframe_fatora, quantity_entry, phpr_entry, in_canvas_rightframe_fatora, total_label, phname_entry_get,\
        df_fat, sz_fat, canvas_rightframe_fatora,\
        canvas_downframe_fatora, in_canvas_downframe_fatora, count_label, fatora

    sz_fat = 0
    df_fat = pd.DataFrame(columns=['index', 'Quantity', 'PhPrice'])

    phname_entry_get = phname_entry.get()
    if not phname_entry_get:
        phname_entry_get = '-----------------'
        answer = tkrmsg.askquestion('تحذير', "لم يتم ادخال اسم الصيدلية!\n متأكد انك تريد المتابعة ؟")
        if answer == 'no':
            return

    date_ph.destroy()
    fatora = tkr.Tk()

    topframe_fatora = tkr.Frame(fatora)
    downframe_fatora = tkr.Frame(fatora)
    rightframe_fatora = tkr.Frame(fatora)

    topframe_fatora.grid(row=0, column=0)
    rightframe_fatora.grid(row=0, column=1, rowspan=2)
    downframe_fatora.grid(row=1)

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
    bupr_label = tkr.Label(topframe_fatora, text='سعر الشراء  ', font=('Arial', '15'))
    quantity_label = tkr.Label(topframe_fatora, text='  الكمية  ', font=('Arial', '15'))
    phpr_label = tkr.Label(topframe_fatora, text='  سعر الصيدلي  ', font=('Arial', '15'))

    bupr_label.grid(row=0, column=0)
    name_label.grid(row=0, column=1)
    cpr_label.grid(row=0, column=2)
    quantity_label.grid(row=0, column=3)
    phpr_label.grid(row=0, column=4)

    cpr_entry_fatora = tkr.Entry(topframe_fatora, font=('Arial', '20'), width=4, textvariable=cur_cpr)
    name_entry_fatora = tkr.Entry(topframe_fatora, font=('Arial', '20'), width=25, textvariable=cur_name)
    bupr_entry_fatora = tkr.Entry(topframe_fatora, font=('Arial', '20'), width=4, textvariable=cur_bupr)
    quantity_entry = tkr.Entry(topframe_fatora, font=('Arial', '20'), width=3)
    phpr_entry = tkr.Entry(topframe_fatora, font=('Arial', '20'), width=4)

    bupr_entry_fatora.grid(row=1, column=0)
    name_entry_fatora.grid(row=1, column=1)
    cpr_entry_fatora.grid(row=1, column=2)
    quantity_entry.grid(row=1, column=3)
    phpr_entry.grid(row=1, column=4)

    count_label = tkr.Label(in_canvas_rightframe_fatora, text='0', font=('Arial', '15'))
    countname_label = tkr.Label(in_canvas_rightframe_fatora, text='  عدد الأصناف  ', font=('Arial', '15'))
    total_label = tkr.Label(in_canvas_rightframe_fatora, text='0.00', font=('Arial', '15'))
    totalname_label = tkr.Label(in_canvas_rightframe_fatora, text='  الإجمالى  ', font=('Arial', '15'))
    save_fatora_button = tkr.Button(in_canvas_rightframe_fatora, text='حفظ', font=('Arial', '20'), width=5,
                                    command=save_pressed)

    total_label.grid(row=0, column=0)
    totalname_label.grid(row=0, column=1)
    count_label.grid(row=0, column=2)
    countname_label.grid(row=0, column=3)
    save_fatora_button.grid(row=0, column=4)

    name1_label = tkr.Label(in_canvas_rightframe_fatora, text='  أسم الصنف  ', font=('Arial', '15'))
    phpr1_label = tkr.Label(in_canvas_rightframe_fatora, text='  سعر القطعة  ', font=('Arial', '15'))
    quantity1_label = tkr.Label(in_canvas_rightframe_fatora, text='  الكمية  ', font=('Arial', '15'))
    phprtotal1_label = tkr.Label(in_canvas_rightframe_fatora, text='  السعر الكلي  ', font=('Arial', '15'))

    phprtotal1_label.grid(row=1, column=0)
    phpr1_label.grid(row=1, column=1)
    quantity1_label.grid(row=1, column=2)
    name1_label.grid(row=1, column=3)

    update_edit('', '', '', 1)

    fatora.mainloop()


def addinfatora_pressed(name):
    global sz_fat
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
    found = get_index(name)
    phprprice = float(phprprice)

    df_fat.loc[sz_fat] = [found, quan, phprprice]
    sz_fat += 1

    show_df_fat()


def removefatora_pressed(idx):
    global sz_fat
    answer = tkrmsg.askquestion("تأكيد", "حذف " + data['Name'][int(df_fat['index'][idx])] + '\n هل انت متأكد ؟')
    if answer == 'no':
        return

    for i in range(idx+1, sz_fat):
        df_fat.loc[i-1] = df_fat.loc[i]
    sz_fat -= 1
    df_fat.drop(df_fat.index[sz_fat], inplace=True)
    show_df_fat()


def show_df_fat():
    total = 0

    for thing in in_canvas_rightframe_fatora.grid_slaves():
        if int(thing.grid_info()["row"]) > 1:
            thing.grid_forget()

    for i in range(sz_fat):
        new_pr = float(df_fat['PhPrice'][i])
        qua = int(df_fat['Quantity'][i])
        add = new_pr*qua
        name2_label = tkr.Label(in_canvas_rightframe_fatora, text=data['Name'][int(df_fat['index'][i])],
                                font=('Arial', '15'), width='15')
        phpr2_label = tkr.Label(in_canvas_rightframe_fatora, text=("%.2f" % new_pr), font=('Arial', '15'))
        quantity2_label = tkr.Label(in_canvas_rightframe_fatora, text=qua, font=('Arial', '15'))
        phprtotal2_label = tkr.Label(in_canvas_rightframe_fatora, text=("%.2f" % add), font=('Arial', '15'))
        removefatora_button = tkr.Button(in_canvas_rightframe_fatora, text='حذف',
                                         font=('Arial', '20'), command=lambda m=i: removefatora_pressed(m), width=3)
        total += add

        phprtotal2_label.grid(row=2+i, column=0)
        phpr2_label.grid(row=2+i, column=1)
        quantity2_label.grid(row=2+i, column=2)
        name2_label.grid(row=2+i, column=3)
        removefatora_button.grid(row=2+i, column=4)
    total_label['text'] = ("%.2f" % total)
    count_label['text'] = len(df_fat)


def save_pressed():
    global dire
    if not sz_fat:
        tkrmsg.showerror('خطأ', 'لم يتم إدخال اى صنف\nلا يمكن طباعة الفاتورة فارغة')
        return

    ans = tkrmsg.askquestion('تأكيد', 'هل انت متأكد من انك تريد حفظ الفاتورة للطباعة ؟')
    if ans == 'no':
        return

    counterfile = open("Database/counter.txt", "r")
    fatora_counter = 1 + int(counterfile.read())
    counterfile.close()
    counterfile = open("Database/counter.txt", "w")
    counterfile.write(str(fatora_counter))
    counterfile.close()

    dire = 'Reports\\' + phname_entry_get + '\\' + str(date.today()) + '\\' + str(fatora_counter) + '\\'
    if not os.path.exists(dire):
        os.makedirs(dire)

    total = 0
    bord = Border(left=Side(style='medium'), right=Side(style='medium'),
                  top=Side(style='medium'), bottom=Side(style='medium'))

    for page in range(1, (sz_fat + 49)//25):
        template = openpyxl.load_workbook('Database/template.xlsx')
        sheet = template['Sheet1']
        sheet['A1'] = 'صفحة ' + str(page)
        sheet['A4'] = date.today()
        sheet['E4'] = 'بيان تسليم'
        sheet['E6'] = 'البضاعة مسلمة إلي : ' + phname_entry_get

        for i in range((page - 1)*25, min(page * 25, sz_fat)):
            idx = int(df_fat['index'][i])
            qua = int(df_fat['Quantity'][i])
            fatpr = float(df_fat['PhPrice'][i])
            add = qua*fatpr
            total += add

            name = data['Name'][idx]
            cupr = data['CPrice'][idx]

            sheet['A' + str(10 + (i % 25))] = ("%.2f" % add)
            sheet['B' + str(10 + (i % 25))] = ("%.2f" % fatpr)
            sheet['C' + str(10 + (i % 25))] = qua//12
            sheet['D' + str(10 + (i % 25))] = qua % 12
            sheet['E' + str(10 + (i % 25))] = name
            sheet['F' + str(10 + (i % 25))] = ("%.2f" % cupr)
            sheet['G' + str(10 + (i % 25))] = i+1

        if (page + 1) == ((sz_fat + 49)//25):
            if page == 1:
                sheet['A37'] = 'تم استلام البضاعة المدونة في صفحة واحدة وقيمتها'
            elif page == 2:
                sheet['A37'] = 'تم استلام البضاعة المدونة في صفحتين وقيمتها'
            else:
                sheet['A37'] = 'تم استلام البضاعة المدونة في ' + str(page) + ' صفحات وقيمتها'
            sheet['A35'] = ("%.2f" % total)
            sheet['F35'] = 'الإجمالى'
            sheet['A35'].border = bord
            sheet['F35'].border = bord
        template.save(dire + str(page) + '.xlsx')
        subprocess.Popen(r'explorer /open,"' + dire + str(page) + '.xlsx')

    save_records()
    fatora.destroy()
    # subprocess.Popen('explorer "C:\path\of\folder"')
    # subprocess.call('explorer ' + dire, shell=True)
    # subprocess.Popen(r'explorer /select,"' + dire + '1.xlsx"')


def save_records():
    os.makedirs(dire + 'DataFrames')
    writer = pd.ExcelWriter(dire + 'DataFrames\\all.xlsx', engine='xlsxwriter')
    data.to_excel(writer, sheet_name='Sheet1')
    writer.save()

    writer = pd.ExcelWriter(dire + 'DataFrames\\fatora.xlsx', engine='xlsxwriter')
    df_fat.to_excel(writer, sheet_name='Sheet1')
    writer.save()


main()
