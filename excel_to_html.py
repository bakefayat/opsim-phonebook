import pandas as pd
from typing import List, Tuple

from bs4 import BeautifulSoup


def separate_columns() -> Tuple[List[str], List[str], List[str], List[str]]:
    """
    Separate columns from an Excel file and return the extracted data
    as separate lists.

    Returns:
        Tuple[List[str], List[str], List[str], List[str]]:
        A tuple containing four lists:
        - unit: A list of strings representing the 'واحد' column data.
        - position: A list of strings representing the 'سمت' column data.
        - name: A list of strings representing the 'نام' column data.
        - phone: A list of strings representing the 'داخلی' column data.
    """
    pan = pd.read_excel('phones.xlsx', header=1)
    df = pd.DataFrame(pan)
    # Fill merged units.
    df['واحد'].fillna(method='ffill', inplace=True)
    unit = df['واحد'].tolist()
    position = df['سمت'].fillna('').tolist()
    name = df['نام خانوادگی'].fillna('').tolist()
    phone = df['شماره داخلی'].fillna('').to_list()
    return (unit, position, name, phone)


def create_zip_from_columns(separated:
                            Tuple[List[str], List[str], List[str], List[str]]):
    """
    Create a zip object from separate lists of columns.

    Args:
        separated: A tuple containing four lists:
        - unit: A list of strings representing the 'unit' column data.
        - position: A list of strings representing the 'position' column data.
        - name: A list of strings representing the 'name' column data.
        - phone: A list of strings representing the 'phone' column data.

    Returns:
        info: A zip object containing tuples with values from the input lists.
    """
    unit, position, name, phone = separated
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


def zip_into_html(zipped: List[Tuple[str, str, str, str]]) -> str:
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
            f"<td class='sahelbold'>{unit}</td>"
            f"<td>{position}</td>"
            f"<td>{name}</td>"
            f"<td>{phone}</td>"
            f"</tr>"
        )
        html_rows += html_row
    return html_rows


def write_into_html(inn_html: str, edition: str) -> None:
    """
    Write new content into an HTML file.

    Args:
        inn_html (str): The new content to be inserted, in HTML format.
        edition (str): The date of new edition.
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
        edition = f'نسخه {edition}'
        bs4_edition = BeautifulSoup(edition, features="html.parser")
        edition_element.extend(bs4_edition.contents)
        # Write into HTML file.
        with open('output.html', 'w', encoding='utf-8') as file:
            file.write(soup.prettify())
            print('Done!')
    else:
        print("One of the target elements not found in HTML.")


if __name__ == '__main__':
    try:
        separated_cols = separate_columns()
        zip_obj = create_zip_from_columns(separated_cols)
        inn_html = zip_into_html(zip_obj)
        edition_input = input('Enter edition like: 1402/01/21 ')
        write_into_html(inn_html, edition_input)

    except FileNotFoundError:
        print('ERR: phones.xlsx file not found.')
