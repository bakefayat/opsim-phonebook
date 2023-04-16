import numpy as np
import pandas as pd


def add_to_set():
    pan = pd.read_excel('phones.xlsx', header=1)
    df = pd.DataFrame(pan)
    # fill merged units.
    df['واحد'].fillna(method='ffill', inplace=True)
    unit = df['واحد'].tolist()
    position = df['سمت'].fillna('').tolist()
    name = df['نام'].fillna('').tolist()
    phone = df['داخلی'].fillna('').to_list()
    units = []
    positions = []
    names = []
    phones = []
    for i in range(0, len(phone)):
        if phone[i] != '':
            names.append(name[i])
            phones.append(phone[i])
            units.append(unit[i])
            positions.append(position[i])
    info = zip(units, positions, names, phones)

    print(list(info))
add_to_set()
