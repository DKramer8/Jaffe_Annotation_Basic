import pandas as pd
import xml.etree.ElementTree as ET
import re
import calendar
from datetime import datetime

month_map = {
    'Jan': '01', 'Januar': '01', 'Janr': '01', 'Ian': '01',
    'Feb': '02', 'Febr': '02', 'Februar': '02',
    'Mar': '03', 'März': '03', 'Maerz': '03',
    'Apr': '04', 'April': '04',
    'May': '05', 'Mai': '05',
    'Jun': '06', 'Juni': '06',
    'Jul': '07', 'Juli': '07',
    'Aug': '08', 'August': '08',
    'Sep': '09', 'Sept': '09', 'September': '09',
    'Oct': '10', 'Okt': '10', 'Oktober': '10',
    'Nov': '11', 'November': '11',
    'Dec': '12', 'Dez': '12', 'Dezember': '12'
}

df = pd.read_excel('Jaffe_Korrekturdatei_I.xlsx')
print(list(df.columns.values))

def get_timespan(row):
    """
    Extracts and constructs a timespan (notBefore and notAfter) based on the given row data.
    This function processes day, month, and year values to generate a valid date range.
    It ensures proper formatting and handles edge cases such as missing or invalid dates.
    Args:
        row (dict): A dictionary or data structure containing the row data from which
                    the timespan is to be extracted.
    Returns:
        tuple: A tuple containing two strings:
               - notBefore (str): The start date of the timespan in "YYYY-MM-DD" format,
                                  or None if the date is invalid.
               - notAfter (str): The end date of the timespan in "YYYY-MM-DD" format,
                                 or None if the date is invalid.
    Notes:
        - If the day is missing or invalid, defaults to the first day of the month for notBefore
          and the last day of the month for notAfter.
        - If the month is not found in the mapping, defaults to January for notBefore and
          December for notAfter.
        - Prints error messages for invalid date formats or other exceptions.
    """
    def check_for_timespan(columnName):        
        """
        Extracts and processes a timespan or single time value from a specified column in a row.
        This function checks the value in the specified column of a row to determine if it contains
        a timespan (e.g., "10-20" or "10—20") or a single time value. It then extracts the numeric
        components of the timespan or single value, removing any non-digit characters.
        Args:
            columnName (str): The name of the column to check in the row.
        Returns:
            tuple: A tuple containing two strings:
                - `first`: The numeric representation of the first time value in the timespan, or the single time value.
                - `last`: The numeric representation of the second time value in the timespan, or the single time value.
        """
        if pd.isna(row[columnName]) or row[columnName] == '': # No entry
            first, last = '', ''
        elif any(c in str(row[columnName]) for c in ['-', '—']): # Timespan
            parts = re.split(r'[-—]', str(row[columnName]), maxsplit=1)
            first, last = parts[0], parts[1]
        else: # No timespan
            first, last = str(row[columnName]), str(row[columnName])
        first, last = re.sub(r'\D', '', first), re.sub(r'\D', '', last) # Remove non-digit characters

        return first, last
       
    firstDay, lastDay = check_for_timespan('Day')
    firstDay, lastDay = firstDay.zfill(2), lastDay.zfill(2) # Ensure two digits for day
    firstMonth, lastMonth = check_for_timespan('Month')
    firstMonth, lastMonth = month_map.get(firstMonth[:4], None), month_map.get(lastMonth[:4], None)
    if firstMonth is None or lastMonth is None:
        firstMonth, lastMonth = '01', '12' # Default to January and December if month is not found in the map
    firstYear, lastYear = check_for_timespan('Year')

    if firstDay == '00': # Day is not found
        firstDay = '01'    
        try:
            _, last_day_of_month = calendar.monthrange(int(lastYear), int(lastMonth))
            lastDay = str(last_day_of_month).zfill(2)    
        except ValueError as e:
            print(e)
    
    try:
        notBefore = f"{firstYear}-{firstMonth}-{firstDay}"
        datetime.strptime(notBefore, "%Y-%m-%d")
        notAfter = f"{lastYear}-{lastMonth}-{lastDay}"
        datetime.strptime(notAfter, "%Y-%m-%d")
    except ValueError as e:
        print(e)
        notBefore, notAfter = None, None
    
    return notBefore, notAfter

for id, row in df.iterrows():
    root = ET.Element('data')
    nr = ''
    row_element = ET.SubElement(root, 'row')
    for col, value in row.items():
        if any(w in col for w in ['pope', 'date', 'place', 'number', 'abstract', 'incipit']): # Transform only the columns we want into xml
            col_element = ET.SubElement(row_element, col)
            if str(value) == 'nan': # Handle NaN values
                value = ''
            col_element.text = str(value)
            if col == 'number':
                nr = str(value)
                nr = nr.replace('?', '')

    notBefore, notAfter = get_timespan(row)
    if notBefore or notAfter: # Only add notBefore and notAfter attributes if they are not None
        notBeforeElement = ET.SubElement(row_element, 'notBefore')
        notBeforeElement.text = str(notBefore)
        notAfterElement = ET.SubElement(row_element, 'notAfter')
        notAfterElement.text = str(notAfter)

    tree = ET.ElementTree(root)
    tree.write(f'transformInput/{id}_{nr}.xml', encoding='utf-8', xml_declaration=True)