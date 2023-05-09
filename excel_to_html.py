import os
import webbrowser

import pandas as pd
from typing import List, Tuple

from bs4 import BeautifulSoup


def separate_columns(file_path: str, header: int, *args):
    """
    Separate columns from an Excel file and return the extracted data
    as separate lists.
    """
    pan = pd.read_excel(file_path, header=header)
    df = pd.DataFrame(pan)
    columns = []
    for col in args:
        column = df[col.get('title')]
        if col.get('fill'):
            column.fillna(method='ffill', inplace=True)
        else:
            column = column.fillna('')

        column.tolist()
        columns.append(column)

    return columns


def create_zip_from_columns(separated):
    """
    Create a zip object from separate lists of columns.

    Args:
        separated: A list containing four lists:
        - unit: A list of strings representing the 'unit' column data.
        - position: A list of strings representing the 'position' column data.
        - name: A list of strings representing the 'name' column data.
        - phone: A list of strings representing the 'phone' column data.

    Returns:
        info: A zip object containing tuples with values from the input lists.
    """
    unit, position, name, phone = separated[:]
    units, positions, names, phones = [], [], [], []
    # Iterate all of valid phones
    for i in range(0, len(phone)):
        if phone[i] != '':
            names.append(name[i])
            phones.append(phone[i])
            # units that have unusual space fixed.
            units.append(unit[i].replace('ـ', ''))
            positions.append(position[i])
    info = zip(units, positions, names, phones)
    return info


def zip_to_html(zipped: List[Tuple[str, str, str, str]]) -> str:
    """
    Generate HTML table rows from a list of tuples.

    Args:
        zipped (List[Tuple[str, str, str, str]]):
        A list of tuples containing four string elements:
        unit, position, name, and phone.

    Returns:
        str: A string containing the generated HTML table rows.
    """
    html_rows = ''
    # Iterate through the zip object
    for unit, position, name, phone in zipped:
        # Generate HTML table row
        html_row = (
            f"<tr>"
            f"<td class='sahel-font'>{unit}</td>"
            f"<td>{position}</td>"
            f"<td>{name}</td>"
            f"<td>{int(phone)}</td>"
            f"</tr>"
        )
        html_rows += html_row
    return html_rows


def get_data():
    """
    get data from user and convert them into full footer html and rtl edition.
    """
    edition_input = input('Enter edition like: 1402/01/21 ')
    footer_html = html_footer(edition_input)
    rtl_edition = reverse_edition(edition_input)

    return footer_html, rtl_edition


def reverse_edition(edition_input: str) -> str:
    """
    Reverse edition from type of YYYY/MM/DD to DD-MM-YYYY
    to show correctly in rtl direction.
    """
    splitted = edition_input.split('/')
    reverse = f'{splitted[2]}-{splitted[1]}-{splitted[0]}'
    return reverse


def html_footer(slash_edition: str) -> str:
    """
    convert edition and url to full footer html text.
    """
    full_text = (
        f'نسخه {slash_edition}<br>'
    )
    return full_text


def get_output_path(dash_edition):
    output_name = f'output/تلفن داخلی سایت سنگان ویرایش {dash_edition}.html'
    full_path = os.path.dirname(os.path.abspath(__file__)) + '/' + output_name
    return full_path


def write_into_html(inn_html: str, footer: str, dash_edition: str) -> None:
    """
    Write new content into an HTML file.

    Args:
        inn_html (str): The new content to be inserted, in HTML format.
        footer (str): containing full context of footer.
        dash_edition (str): edition of file appending on output file name.
    """
    with open('base.html', 'r', encoding='utf-8') as file:
        html_string = file.read()

    soup = BeautifulSoup(html_string, 'html.parser')
    tbody_target_element = soup.select_one('.target-id')
    edition_element = soup.select_one('.edition')

    if tbody_target_element and edition_element:
        # Insert new content into the tbody target element
        bs4_html = BeautifulSoup(inn_html, 'html.parser')
        tbody_target_element.extend(bs4_html.contents)
        # Insert new content into the edition element
        bs4_edition = BeautifulSoup(footer, features="html.parser")
        edition_element.extend(bs4_edition.contents)
        # Write into HTML file.
        output_path = get_output_path(dash_edition)
        with open(output_path, 'w', encoding='utf-8') as file:
            file.write(soup.prettify())
            print('Done!')
            webbrowser.open(output_path)
    else:
        print("One of the target elements wasn't found in HTML.")


if __name__ == '__main__':
    try:
        # seprate columns of file with starting point.
        col1 = {'title': 'واحد', 'format': str, 'fill': True}
        col2 = {'title': 'سمت', 'format': str, 'fill': True}
        col3 = {'title': 'نام خانوادگی', 'format': str, 'fill': False}
        col4 = {'title': 'شماره داخلی', 'format': int, 'fill': False}
        separated_cols = separate_columns('phones.xlsx', 1, col1, col2, col3, col4)
        zip_obj = create_zip_from_columns(separated_cols)
        rows_context = zip_to_html(zip_obj)
        footer_context, edition = get_data()
        write_into_html(rows_context, footer_context, edition)

    except FileNotFoundError:
        print('ERR: phones.xlsx file not found.')

    except IndexError:
        print('ERR: Date format is incorrect')

    except KeyError:
        print('ERR: Check Columns head values or startpoint.')

    except KeyboardInterrupt:
        print('ERR: you ended the program.')
