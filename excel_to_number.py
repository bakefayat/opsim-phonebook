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
            # phones changed from float type to int
            phones.append(int(phone[i]))
            # units that has unusual space fixed.
            units.append(unit[i].replace('ـ', ''))
            positions.append(position[i])
    info = zip(units, positions, names, phones)
    
    html_rows = ''
    # Iterate through the zip object
    for unit, position, name, phone in info:
        # Generate HTML table row
        html_row = f"<tr>\n\t<td>{unit}</td>\n\t<td>{position}</td>\n\t<td>{name}</td>\n\t<td>{phone}</td>\n</tr>\n"
        html_rows += html_row

    # Write to HTML file.
    with open('output.html', 'w', encoding='utf-8') as f:
        f.write(html_rows)

add_to_set()
