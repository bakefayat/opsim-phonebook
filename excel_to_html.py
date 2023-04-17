import pandas as pd


def excel_to_zip():
    pan = pd.read_excel('phones.xlsx', header=1)
    df = pd.DataFrame(pan)
    # Fill merged units.
    df['واحد'].fillna(method='ffill', inplace=True)
    unit = df['واحد'].tolist()
    position = df['سمت'].fillna('').tolist()
    name = df['نام'].fillna('').tolist()
    phone = df['داخلی'].fillna('').to_list()
    units, positions, names, phones = [], [], [], []
    # Iterate all of valid phones
    for i in range(0, len(phone)):
        if phone[i] != '':
            names.append(name[i])
            # phones changed from float type to int
            phones.append(int(phone[i]))
            # units that has unusual space fixed.
            units.append(unit[i].replace('ـ', ''))
            positions.append(position[i])
    info = zip(units, positions, names, phones)
    return info


def zip_into_html(zipped):
    html_rows = ''
    # Iterate through the zip object
    for unit, position, name, phone in zipped:
        # Generate HTML table row
        html_row = (
            f"<tr>\n"
            f"\t<td>{unit}</td>\n"
            f"\t<td>{position}</td>\n"
            f"\t<td>{name}</td>\n"
            f"\t<td>{phone}</td>\n"
            f"</tr>\n"
        )
        html_rows += html_row
    return html_rows


def write_to_html(rows):
    # Write to HTML file.
    with open('output.html', 'w', encoding='utf-8') as f:
        f.write(rows)
        print('Done! now use output.html file')


try:
    zip_obj = excel_to_zip()
    inn_html = zip_into_html(zip_obj)
    write_to_html(inn_html)
except FileNotFoundError:
    print('ERR: phones.xlsx file not found.')
