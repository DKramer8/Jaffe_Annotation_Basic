import pandas as pd
import xml.etree.ElementTree as ET
import re

df = pd.read_excel('Jaffe_Korrekturdatei_I.xlsx')
print(list(df.columns.values))

def get_timespan(row):
    def check_for_timespan(columnName):        
        if pd.isna(row[columnName]) or row[columnName] == '': # No entry
            first, last = '', ''
        elif any(c in str(row[columnName]) for c in ['-', '—']): # Entry that is a timespan
            parts = re.split(r'[-—]', str(row[columnName]), maxsplit=1)
            first, last = parts[0], parts[1]
        else: # Entry that is not a timespan
            first, last = str(row[columnName]), str(row[columnName])
        
        first, last = first.replace(' ', ''), last.replace(' ', '') # Delete unneccassary whitespace
        return first, last
    
    firstYear, lastYear = check_for_timespan('Year')
    firstMonth, lastMonth = check_for_timespan('Month')
    if firstMonth != '': # If there is a month, add '-' before to include it in whole date
        firstMonth, lastMonth = '-' + firstMonth, '-' + lastMonth
    firstDay, lastDay = check_for_timespan('Day')
    if firstDay != '': # If there is a day, add '-' before to include it in whole date
        firstDay, lastDay = '-' + firstDay, '-' + lastDay
    
    notBefore = firstYear + firstMonth + firstDay
    notAfter = lastYear + lastMonth + lastDay
    return notBefore, notAfter

root = ET.Element('data')
for _, row in df.iterrows():
    row_element = ET.SubElement(root, 'row')
    for col, value in row.items():
        if any(w in col for w in ['pope', 'date', 'place', 'number', 'abstract', 'incipit']): # Transform only the columns we want into xml
            col_element = ET.SubElement(row_element, col)
            col_element.text = str(value)

    notBefore, notAfter = get_timespan(row)
    notBeforeElement = ET.SubElement(row_element, 'notBefore')
    notBeforeElement.text = str(notBefore)
    notAfterElement = ET.SubElement(row_element, 'notAfter')
    notAfterElement.text = str(notAfter)

tree = ET.ElementTree(root)
tree.write('output.xml', encoding='utf-8', xml_declaration=True)