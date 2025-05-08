import pandas as pd
import xml.etree.ElementTree as ET
import re
import calendar
from datetime import datetime

MONTH_DICT = {
    'januar': '01', 'janr': '01', 'ianuar': '01', 'ianr': '01',
    'februar': '02',
    'märz': '03', 'mart': '03', 'maerz': '03',
    'april': '04',
    'may': '05', 'mai': '05',
    'iuni': '06', 'juni': '06',
    'iuli': '07', 'juli': '07',
    'august': '08',
    'september': '09',
    'october': '10', 'oktober': '10',
    'november': '11',
    'december': '12', 'dezember': '12'
}
COLUMNS = ['pope', 'date', 'place', 'number', 'abstract', 'incipit']
INPUT_FILE = 'Jaffe_Korrekturdatei_I.xlsx'

df = pd.read_excel(INPUT_FILE)
#df = pd.read_excel('jaffe_test.xlsx')
print(list(df.columns.values))

def get_dates(row, nr):
    """
    Extracts and constructs a timespan (notBefore and notAfter) based on the given row data.
    This function processes day, month, and year values to generate a valid date range.
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
        a timespan (e.g., "10-20", "10—20" or "10 20") or a single time value. It then extracts the numeric
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
        elif any(c in str(row[columnName]) for c in ['-', '—']): # Timespan divided by '-' or '—'
            parts = re.split(r'[-—]', str(row[columnName]), maxsplit=1)
            first, last = parts[0], parts[1]
        elif re.search(r'(?<=\d)\s(?=\d)|(?<!\d)\s(?!\d)|(?<=\d\.)\s(?=\d)|(?<=\D\.)\s(?=\D)', str(row[columnName])): # Timespan divided by whitespace
            parts = re.split(r'\s', str(row[columnName]), maxsplit=1)
            first, last = parts[0], parts[1]
        else: # No timespan
            first, last = str(row[columnName]), str(row[columnName])

        return first, last
       
    def get_month_digits(month_str):
        """
        Converts a month name or abbreviation to its corresponding digits.
        Args:
            month_str (str): The name or abbreviation of the month (e.g., "January", "Jan", "Feb").
        Returns:
            int: The digits corresponding to the month (01 for January, 02 for February, etc.),
            None if the input does not match any month in the MONTH_DICT.
        """
        
        for key, value in MONTH_DICT.items():
            if re.search(month_str.lower(), key):
                return value
        return '' 

    firstYear, lastYear = check_for_timespan('Year')
    firstYear, lastYear = re.sub(r'\D', '', firstYear), re.sub(r'\D', '', lastYear) # Remove non-digit characters

    firstMonth, lastMonth = check_for_timespan('Month')
    firstMonth, lastMonth = re.sub(r'[^a-zA-ZäöüÄÖÜß]', '', firstMonth), re.sub(r'[^a-zA-ZäöüÄÖÜß]', '', lastMonth) # Remove non-alphabetic characters
    if firstMonth != '' or lastMonth != '':
        firstMonth, lastMonth = get_month_digits(firstMonth), get_month_digits(lastMonth)
    else:
        firstMonth, lastMonth = '01', '12' # If month is not found -> get full year timespan
    
    firstDay, lastDay = check_for_timespan('Day')
    firstDay, lastDay = re.sub(r'\D', '', firstDay), re.sub(r'\D', '', lastDay) # Remove non-digit characters
    firstDay, lastDay = firstDay.zfill(2), lastDay.zfill(2) # Ensure two digits for day
    if firstDay == '00': # Happens if day is not found
        firstDay = '01'  # Day is not found -> get full month timespan  
        try:
            _, last_day_of_month = calendar.monthrange(int(lastYear), int(lastMonth))
            lastDay = str(last_day_of_month).zfill(2)    
        except ValueError as e:
            print(f'{nr}, {row['JL']}, {row['Year']}.{row['Month']}.{row['Day']}: Cant get first and last day of month: {e}')
    
    try:
        notBefore = f"{firstYear}-{firstMonth}-{firstDay}"
        datetime.strptime(notBefore, "%Y-%m-%d")
        notAfter = f"{lastYear}-{lastMonth}-{lastDay}"
        datetime.strptime(notAfter, "%Y-%m-%d")
    except ValueError as e:
        print(f'{nr}, {row['JL']}, {row['Year']}.{row['Month']}.{row['Day']}: Cant build full date: {e}')
        notBefore, notAfter = None, None
    
    return notBefore, notAfter

root = ET.Element('data')
for id, row in df.iterrows():
    #root = ET.Element('data')
    nr = ''
    row_element = ET.SubElement(root, 'row')
    for col, value in row.items():
        if any(w in col for w in COLUMNS): # Transform only the columns we want into xml
            col_element = ET.SubElement(row_element, col)
            if str(value) == 'nan': # Handle NaN values
                value = ''
            col_element.text = str(value)
            if col == 'number':
                nr = str(value)
                nr = nr.replace('?', '')

    if pd.isna(row['date']):
        pass
    else:
        notBefore, notAfter = get_dates(row, id)
        if notBefore or notAfter: # Only add notBefore and notAfter attributes if they are not None
            notBeforeElement = ET.SubElement(row_element, 'notBefore')
            notBeforeElement.text = str(notBefore)
            notAfterElement = ET.SubElement(row_element, 'notAfter')
            notAfterElement.text = str(notAfter)

    tree = ET.ElementTree(root)
    #tree.write(f'transformInput/{id}_{nr}.xml', encoding='utf-8', xml_declaration=True)
tree.write(f'test/{id}_{nr}.xml', encoding='utf-8', xml_declaration=True)