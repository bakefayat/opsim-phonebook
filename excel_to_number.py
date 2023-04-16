import numpy as np
import pandas as pd


def add_to_set():
    pan = pd.read_excel('phones.xlsx', header=1)
    cell_value = pan.at[2, 'داخلی']
    df = pd.DataFrame(pan)
    df['واحد'].fillna(method='ffill', inplace=True)

    unit = df['واحد'].tolist()
    position = df['سمت'].tolist()
    name = df['نام'].tolist()
    phone = df['داخلی'].tolist()
    units = []
    positions = []
    names = []
    phones = []
    for i in range(0, len(phone)):
        if phone != 'nan':
            names.append(name[i])
            phones.append(phone[i])
            units.append(unit[i])
            positions.append(position[i])
    info = zip(units, positions, names, phones)
    print(list(info))


add_to_set()