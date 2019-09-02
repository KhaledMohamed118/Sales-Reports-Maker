import os
import pandas as pd


def save_database():
    writer = pd.ExcelWriter(Database_Name, engine='xlsxwriter')
    data.to_excel(writer, sheet_name='Sheet1')
    writer.save()

    colfile2 = open("Database/column.txt", "w")
    colfile2.write(str(colchar))
    colfile2.close()


LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
Database_Name = "Database/Book.xlsx"

colfile = open("Database/column.txt", "r")
colchar = int(colfile.read())
colfile.close()

if colchar < 26:
    tempcols = "B:" + LETTERS[colchar]
else:
    tempcols = "B:" + LETTERS[(colchar // 26) - 1] + LETTERS[colchar % 26]

data = pd.read_excel(Database_Name, usecols=tempcols)

Database_Size = len(data)

for filename in os.listdir("Reports"):
    print(filename)
    if filename not in data.columns:
        listofzeros = [0.0] * Database_Size
        data[filename] = listofzeros
        colchar += 1

    for subdir, dirs, files in os.walk(os.path.join("Reports", filename)):
        for file in files:
            if file == 'fatora.xlsx':
                cur_fatora = os.path.join(subdir, file)
                print(cur_fatora)
                cur_fatora = pd.read_excel(cur_fatora, usecols="B:D")
                if "Name" not in cur_fatora.columns:
                    continue
                for j in range(len(cur_fatora)):
                    for i in range(Database_Size):
                        if data["Name"][i] == cur_fatora["Name"][j]:
                            data.at[i, filename] = float(cur_fatora["PhPrice"][j])
                            break
    print("========")

save_database()
