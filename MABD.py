import pandas as pd
import numpy as np
from copy import deepcopy

Database_Name = "Book.xlsx"

data = pd.read_excel(Database_Name)
Database_Size = len(data)


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
        if (name in data["Name"][i]) and (cpr == 0 or cpr == data["CPrice"][i]) and (phpr == 0 or phpr == data["PhPrice"][i]):
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


print(data)
