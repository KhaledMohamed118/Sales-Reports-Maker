import tkinter as tkr
import tkinter.messagebox as tkrmsg
import pandas as pd
import numpy as np
from copy import deepcopy
from datetime import date


def main():
    global data, root, Database_Size, Database_Name
    Database_Name = "Book.xlsx"
    data = pd.read_excel(Database_Name)
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
    canvas_rightframe_fatora.configure(scrollregion=canvas_rightframe_fatora.bbox("all"), width=500, height=500)


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


def search(name, cpr, phpr):
    if notnumber(cpr):
        cpr = 0
    else:
        cpr = '0' + cpr
        if '.' in cpr:
            cpr += '0'
        cpr = float(cpr)

    if notnumber(phpr):
        phpr = 0
    else:
        phpr = '0' + phpr
        if '.' in phpr:
            phpr += '0'
        phpr = float(phpr)

    sz = 0
    df = pd.DataFrame(columns=['Name', 'CPrice', 'PhPrice'])

    for i in range(Database_Size):
        t = data.loc[i]
        if (name.lower() in t["Name"].lower()) and ((not cpr) or (cpr == float(t["CPrice"]))) and ((not phpr) or (phpr == float(t["PhPrice"]))):
            df.loc[sz] = data.loc[i]
            sz += 1

    return df


def showall(event):
    root.destroy()
    showall_window = tkr.Tk()

    cpr_label = tkr.Label(showall_window, text='  سعر المستهلك  ', font=('Arial', '25'))
    name_label = tkr.Label(showall_window, text='  أسم الصنف  ', font=('Arial', '25'))
    phpr_label = tkr.Label(showall_window, text='  سعر الصيدلي ', font=('Arial', '25'))

    phpr_label.grid(row=0, column=0)
    name_label.grid(row=0, column=1)
    cpr_label.grid(row=0, column=2)

    for i in range(Database_Size):
        cpr_label = tkr.Label(showall_window, text=("%.2f" % data['CPrice'][i]), font=('Arial', '15'))
        name_label = tkr.Label(showall_window, text=data['Name'][i], font=('Arial', '15'))
        phpr_label = tkr.Label(showall_window, text=("%.2f" % data['PhPrice'][i]), font=('Arial', '15'))

        phpr_label.grid(row=i+1, column=0)
        name_label.grid(row=i+1, column=1)
        cpr_label.grid(row=i+1, column=2)

    showall_window.mainloop()
    main()


def addone(event):
    root.destroy()
    global cpr_entry_add, name_entry_add, phpr_entry_add
    addone_window = tkr.Tk()

    cpr_label = tkr.Label(addone_window, text='سعر المستهلك', font=('Arial', '15'))
    name_label = tkr.Label(addone_window, text='أسم الصنف', font=('Arial', '15'))
    phpr_label = tkr.Label(addone_window, text='سعر الصيدلي', font=('Arial', '15'))

    cpr_entry_add = tkr.Entry(addone_window, font=('Arial', '20'), width=5)
    name_entry_add = tkr.Entry(addone_window, font=('Arial', '20'), width=30)
    phpr_entry_add = tkr.Entry(addone_window, font=('Arial', '20'), width=5)

    add_button = tkr.Button(addone_window, text='إضافة', font=('Arial', '20'), width=10, command=add_button_pressed)
    # add_button.flash()
    # add_button.bind("<Button-1>", add_button_pressed)

    phpr_label.grid(row=0)
    name_label.grid(row=0, column=1)
    cpr_label.grid(row=0, column=2)

    phpr_entry_add.grid(row=1, column=0)
    name_entry_add.grid(row=1, column=1)
    cpr_entry_add.grid(row=1, column=2)

    add_button.grid(row=2, columnspan=3)

    addone_window.mainloop()
    main()


def add_button_pressed():
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
            tt1 = '0' + cpr_entry_add.get()
            tt2 = '0' + phpr_entry_add.get()
            if '.' in tt1:
                tt1 += '0'
            if '.' in tt2:
                tt2 += '0'
            add_product(name_entry_add.get(), float(tt1), float(tt2))
            tkrmsg.showinfo('تم', 'تم اضافة الصنف بنجاح')


def add_product(name, cpr, phpr):
    global Database_Size
    data.loc[Database_Size] = [name, cpr, phpr]
    Database_Size += 1
    save_database()


def editone(event):
    global downframe_edit
    root.destroy()
    editone_window = tkr.Tk()

    topframe = tkr.Frame(editone_window)
    downframe_edit = tkr.Frame(editone_window)

    topframe.pack()
    downframe_edit.pack()

    cur_name = tkr.StringVar()
    cur_cpr = tkr.StringVar()
    cur_phpr = tkr.StringVar()

    cur_name.trace("w", lambda name, index, mode: update_edit(cur_name.get(), cur_cpr.get(), cur_phpr.get(), 0))
    cur_cpr.trace("w", lambda name, index, mode: update_edit(cur_name.get(), cur_cpr.get(), cur_phpr.get(), 0))
    cur_phpr.trace("w", lambda name, index, mode: update_edit(cur_name.get(), cur_cpr.get(), cur_phpr.get(), 0))

    cpr_label = tkr.Label(topframe, text='  سعر المستهلك  ', font=('Arial', '15'))
    name_label = tkr.Label(topframe, text='  أسم الصنف  ', font=('Arial', '15'))
    phpr_label = tkr.Label(topframe, text='سعر الصيدلي', font=('Arial', '15'))

    phpr_label.grid(row=0, column=0)
    name_label.grid(row=0, column=1)
    cpr_label.grid(row=0, column=2)

    cpr_entry_edit = tkr.Entry(topframe, font=('Arial', '20'), width=5, textvariable=cur_cpr)
    name_entry_edit = tkr.Entry(topframe, font=('Arial', '20'), width=30, textvariable=cur_name)
    phpr_entry_edit = tkr.Entry(topframe, font=('Arial', '20'), width=5, textvariable=cur_phpr)

    phpr_entry_edit.grid(row=1, column=0)
    name_entry_edit.grid(row=1, column=1)
    cpr_entry_edit.grid(row=1, column=2)
    update_edit('', '', '', 0)
    editone_window.mainloop()
    main()


def update_edit(name, cpr, phpr, whocalled):
    if whocalled:
        fr = downframe_fatora
    else:
        fr = downframe_edit

    for i in fr.winfo_children():
        i.destroy()
    df = search(name, cpr, phpr)
    for i in range(len(df)):
        cpr_label = tkr.Label(fr, text=("%.2f" % df['CPrice'][i]), font=('Arial', '15'), width=5)
        name_label = tkr.Label(fr, text=df['Name'][i], font=('Arial', '15'), width=45)
        phpr_label = tkr.Label(fr, text=("%.2f" % df['PhPrice'][i]), font=('Arial', '15'), width=5)

        phpr_label.grid(row=i, column=0)
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
    global cpr_entry_edit, name_entry_edit, phpr_entry_edit, edit_window

    found = get_index(name)

    edit_window = tkr.Tk()

    cpr_label_edit = tkr.Label(edit_window, text='سعر المستهلك', font=('Arial', '15'))
    name_label_edit = tkr.Label(edit_window, text='أسم الصنف', font=('Arial', '15'))
    phpr_label_edit = tkr.Label(edit_window, text='سعر الصيدلي', font=('Arial', '15'))

    cpr_entry_edit = tkr.Entry(edit_window, font=('Arial', '20'), width=5)
    name_entry_edit = tkr.Entry(edit_window, font=('Arial', '20'), width=30)
    phpr_entry_edit = tkr.Entry(edit_window, font=('Arial', '20'), width=5)

    cpr_entry_edit.insert(0, data['CPrice'][found])
    name_entry_edit.insert(0, data['Name'][found])
    phpr_entry_edit.insert(0, data['PhPrice'][found])

    edit_button = tkr.Button(edit_window, text='تعديل', font=('Arial', '20'), width=10,  command=lambda m=found: edit_buttpon_clicked(m))

    phpr_label_edit.grid(row=0)
    name_label_edit.grid(row=0, column=1)
    cpr_label_edit.grid(row=0, column=2)

    phpr_entry_edit.grid(row=1, column=0)
    name_entry_edit.grid(row=1, column=1)
    cpr_entry_edit.grid(row=1, column=2)

    edit_button.grid(row=2, columnspan=3)

    edit_window.mainloop()


def edit_buttpon_clicked(idx):
    if not cpr_entry_edit.get() or not name_entry_edit.get() or not phpr_entry_edit.get():
        tkrmsg.showerror('خطأ', 'لا يمكن ترك احد الخانات فارغة')

    elif (get_index(name_entry_edit.get()) < Database_Size) and (get_index(name_entry_edit.get()) != idx):
        tkrmsg.showerror('خطأ', 'هذا الأسم موجود من قبل')

    else:
        if notnumber(cpr_entry_edit.get()):
            tkrmsg.showerror('خطأ', 'خطأ فى سعر المستهلك')
        elif notnumber(phpr_entry_edit.get()):
            tkrmsg.showerror('خطأ', 'خطأ فى سعر الصيدلي')
        else:
            tt1 = '0' + cpr_entry_edit.get()
            tt2 = '0' + phpr_entry_edit.get()
            if '.' in tt1:
                tt1 += '0'
            if '.' in tt2:
                tt2 += '0'

            data['Name'][idx] = name_entry_edit.get()
            data['CPrice'][idx] = float(tt1)
            data['PhPrice'][idx] = float(tt2)
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


def start_buttpon_clicked():
    global downframe_fatora, quantity_entry, total, in_canvas_rightframe_fatora, total_label, df_fat, sz_fat, canvas_rightframe_fatora
    total = 0
    sz_fat = 0
    df_fat = pd.DataFrame(columns=['index', 'Quantity'])

    phname = phname_entry.get()
    if not phname:
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
    cur_phpr = tkr.StringVar()

    cur_name.trace("w", lambda name, index, mode: update_edit(cur_name.get(), cur_cpr.get(), cur_phpr.get(), 1))
    cur_cpr.trace("w", lambda name, index, mode: update_edit(cur_name.get(), cur_cpr.get(), cur_phpr.get(), 1))
    cur_phpr.trace("w", lambda name, index, mode: update_edit(cur_name.get(), cur_cpr.get(), cur_phpr.get(), 1))

    cpr_label = tkr.Label(topframe_fatora, text='  سعر المستهلك  ', font=('Arial', '15'))
    name_label = tkr.Label(topframe_fatora, text='  أسم الصنف  ', font=('Arial', '15'))
    phpr_label = tkr.Label(topframe_fatora, text='سعر الصيدلي  ', font=('Arial', '15'))
    quantity_label = tkr.Label(topframe_fatora, text='  الكمية  ', font=('Arial', '15'))
    addinfatora_label = tkr.Label(topframe_fatora, text='    ', font=('Arial', '15'))

    phpr_label.grid(row=0, column=0)
    name_label.grid(row=0, column=1)
    cpr_label.grid(row=0, column=2)
    quantity_label.grid(row=0, column=3)
    addinfatora_label.grid(row=0, column=4)

    cpr_entry_fatora = tkr.Entry(topframe_fatora, font=('Arial', '20'), width=5, textvariable=cur_cpr)
    name_entry_fatora = tkr.Entry(topframe_fatora, font=('Arial', '20'), width=30, textvariable=cur_name)
    phpr_entry_fatora = tkr.Entry(topframe_fatora, font=('Arial', '20'), width=5, textvariable=cur_phpr)
    quantity_entry = tkr.Entry(topframe_fatora, font=('Arial', '20'), width=5)

    phpr_entry_fatora.grid(row=1, column=0)
    name_entry_fatora.grid(row=1, column=1)
    cpr_entry_fatora.grid(row=1, column=2)
    quantity_entry.grid(row=1, column=3)

    total_label = tkr.Label(in_canvas_rightframe_fatora, text=("%.2f" % total), font=('Arial', '15'))
    totalname_label = tkr.Label(in_canvas_rightframe_fatora, text='الإجمالى : ', font=('Arial', '15'))

    total_label.grid(row=0, column=0)
    totalname_label.grid(row=0, column=1)

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
    global sz_fat, total
    quan = quantity_entry.get()
    if (notnumber(quan)) or ('.' in quan) or (not quan) or (not int(quan)):
        tkrmsg.showerror("خطأ", 'خطأ فى الكمية .. لم يتم اضافة الصنف إلى الفاتورة')
        return

    quan = int(quan)
    found = get_index(name)
    new_pr = float(data['PhPrice'][found])*quan
    name2_label = tkr.Label(in_canvas_rightframe_fatora, text=name, font=('Arial', '15'), width='15')
    phpr2_label = tkr.Label(in_canvas_rightframe_fatora, text=("%.2f" % data['PhPrice'][found]), font=('Arial', '15'))
    quantity2_label = tkr.Label(in_canvas_rightframe_fatora, text=quan, font=('Arial', '15'))
    phprtotal2_label = tkr.Label(in_canvas_rightframe_fatora, text=("%.2f" % new_pr), font=('Arial', '15'))

    phprtotal2_label.grid(row=2+sz_fat, column=0)
    phpr2_label.grid(row=2+sz_fat, column=1)
    quantity2_label.grid(row=2+sz_fat, column=2)
    name2_label.grid(row=2+sz_fat, column=3)

    df_fat.loc[sz_fat] = [found, quan]
    sz_fat += 1
    total += new_pr
    total_label['text'] = ("%.2f" % total)


def temp(par):
    today = date.today()
    print(today)  # '2017-12-26'


def temp2():
    print('def')


main()
