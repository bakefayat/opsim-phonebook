import pandas as pd
from bs4 import BeautifulSoup


def seprate_columns():
    pan = pd.read_excel('phones.xlsx', header=1)
    df = pd.DataFrame(pan)
    # Fill merged units.
    df['واحد'].fillna(method='ffill', inplace=True)
    unit = df['واحد'].tolist()
    position = df['سمت'].fillna('').tolist()
    name = df['نام'].fillna('').tolist()
    phone = df['داخلی'].fillna('').to_list()
    return (unit, position, name, phone)


def columns_to_zip(seprated):
    unit, position, name, phone = seprated
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


if __name__ == '__main__':
    try:
        seprated_cols = seprate_columns()
        zip_obj = columns_to_zip(seprated_cols)
        inn_html = zip_into_html(zip_obj)
        
        with open('base.html', 'r', encoding='utf-8') as file:
            html_string = file.read()
        
        soup = BeautifulSoup(html_string, 'html.parser')
        target_element = soup.select_one('.target-id')

        if target_element:
            # Insert new content into the target element
            bs4_html = BeautifulSoup(inn_html, 'html.parser')
            target_element.extend(bs4_html.contents)
             
            with open('output.html', 'w', encoding='utf-8') as file:
                file.write(soup.prettify())
        else:
            print("Target element not found in HTML.")

        # write_to_html(inn_html)
    except FileNotFoundError:
        print('ERR: phones.xlsx file not found.')
