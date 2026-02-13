import pandas as pd
import os
import re
import math
from rapidfuzz import process, fuzz

#------------------------------------------------------------
#------------------------------------------------------------
# GLOBALS
#------------------------------------------------------------
#------------------------------------------------------------
DIR_PATH = os.getcwd()
DICTIONARY_FILE = f'{DIR_PATH}/data/static/jaffe_dicts.xlsx'
INPUT_PATH = f'{DIR_PATH}/data/output/'
OUTPUT_FILE = f'{DIR_PATH}/data/postprocessed/postprocessed_'

COLUMNS = ['date', 'pope', 'place', 'number'] 
MIN_LEVDIS_SCORE_NUMBER = 85
MIN_LEVDIS_SCORE_YEAR = 70
MIN_LEVDIS_SCORE_MONTH = 66
MIN_LEVDIS_SCORE_DAY = 70
MIN_LEVDIS_SCORE = 70

#------------------------------------------------------------
#------------------------------------------------------------
# Functions
#------------------------------------------------------------
#------------------------------------------------------------
def replace_by_dict(text: str, dictionary: list, lev_distance: int, row_nmb: int) -> str:
    '''
    Replaces a given text with the closest match from a provided dictionary if the Levenshtein distance score is above a certain threshold.
    
    :param text: The text to be replaced
    :param dictionary: A list of valid entries to compare against
    :param lev_distance: The minimum Levenshtein distance score required for a replacement to occur
    :param row_nmb: The row number in the original DataFrame, used for logging purposes
    :return: The original text if no close match is found, or the closest match from the dictionary if the score is above the threshold
    '''

    match, score, _ = process.extractOne(
        text,
        dictionary,
        scorer=fuzz.ratio
    )
    if score >= lev_distance and score != 100:
        print(f"Ersetze '{text}' durch '{match}' (Score: {score}) at row {row_nmb+2}")
        return match
    else:
        return text
    
def split_date(full_date: str) -> tuple:
    '''
    Splits a date string into its components (year, month, day) and removes any unwanted characters. If any component is missing, it returns a space for that component.
    
    :param full_date: The full date string to be split
    :returns: A tuple containing the year, month, and day components
    '''

    year = full_date[:4]
    month = re.sub(r"[^a-zA-Z()]", "", full_date[4:])
    if month.endswith('()'):
        month = month[:-2]
    day = re.sub(r"[^0-9\[\]]", "", full_date[4:])
    if not year:
        year = ' '
    if not month:
        month = ' '
    if not day:
        day = ' '
    return year, month, day

def clean_number(s):
    if s is None or (isinstance(s, float) and math.isnan(s)):
        return ''
    return re.sub(r'[^0-9() ]', '', str(s))

# def process_number_column(numbers_raw: list) -> list:
#     '''
#     Postprocesses the OCR results of the "number" column based on a strict +1 sequence rule. (Made by ChatGPT)
    
#     :param numbers_raw: List of raw OCR strings from the number column
#     :return: List of corrected values
#     '''
    
#     processed = []
#     last_valid_number = None

#     for idx, value in enumerate(numbers_raw):
#         value_str = str(value).strip()

#         # Wenn leer → 그대로 übernehmen
#         if not value_str or value_str.lower() == 'nan':
#             processed.append('')
#             continue

#         value_str = re.sub(r'[^0-9() ]', '', value_str)

#         # Hauptzahl extrahieren (erste Zahl im String)
#         match = re.search(r'\d+', value_str)
#         if not match:
#             # Keine Zahl gefunden → unverändert lassen
#             processed.append(value_str)
#             continue

#         ocr_number = match.group()
#         rest = value_str[len(ocr_number):]  # Klammerteil etc. bleibt erhalten

#         # Erste gefundene Zahl initialisiert Sequenz
#         if last_valid_number is None:
#             last_valid_number = int(ocr_number)
#             processed.append(value_str)
#             continue

#         expected_number = str(last_valid_number + 1)

#         # Wenn korrekt → übernehmen
#         if ocr_number == expected_number:
#             last_valid_number += 1
#             processed.append(value_str)
#             continue

#         # Wenn nicht korrekt → Ähnlichkeit prüfen
#         score = fuzz.ratio(ocr_number, expected_number)

#         if score >= MIN_LEVDIS_SCORE_NUMBER:
#             print(f"Ersetze '{ocr_number}' durch '{expected_number}' (Score: {score}) at row {idx+2}")
#             corrected_value = expected_number + rest
#             processed.append(corrected_value)
#             last_valid_number += 1
#         else:
#             # OCR-Wert bleibt stehen
#             processed.append(value_str)
#             last_valid_number = int(ocr_number)

#     return processed

#------------------------------------------------------------
#------------------------------------------------------------
# Main
#------------------------------------------------------------
#------------------------------------------------------------
if __name__ == '__main__':
    print('Please enter file name (must end in .xlsx)')
    user_input = input()
    input_file = INPUT_PATH + user_input
    output_file = OUTPUT_FILE + user_input

    df_dict = pd.read_excel(DICTIONARY_FILE)
    df_input = pd.read_excel(input_file)
    output_df = df_input

    for column in COLUMNS:        
        if column == 'date':
            year_dict = df_dict['year'].dropna().astype(str).tolist()
            month_dict = df_dict['month'].dropna().astype(str).tolist()
            month_dict = [x.replace('.', '') for x in month_dict] # Um den String Vergleich mit dem Dict nicht zu erschweren, werden die Punkte aus dem Dict entfernt und später manuell hinzugefügt
            day_dict = df_dict['day'].dropna().astype(int).astype(str).tolist()

            years_raw = [split_date(t)[0] for t in df_input['date'].astype(str)]
            months_raw = [split_date(t)[1] for t in df_input['date'].astype(str)]
            days_raw = [split_date(t)[2] for t in df_input['date'].astype(str)]

            years = [replace_by_dict(y, year_dict, MIN_LEVDIS_SCORE_YEAR, idx) for idx, y in enumerate(years_raw)]
            months = [replace_by_dict(m, month_dict, MIN_LEVDIS_SCORE_MONTH, idx) for idx, m in enumerate(months_raw)]
            months = [x + '.' if x != ' ' else x for x in months]
            days = [replace_by_dict(d, day_dict, MIN_LEVDIS_SCORE_DAY, idx) for idx, d in enumerate(days_raw)]
            days = [x + '.' if x != ' ' else x for x in days]

            i = 0
            while i < len(years):
                output_df.loc[i, 'date'] = f'{years[i]} {months[i]} {days[i]}'
                i += 1

        elif column == 'number':
            numbers_raw = df_input['number'].tolist()
            output_df['number'] = [clean_number(s) for s in numbers_raw]

        else:
            dictionary = df_dict[column].dropna().astype(str).tolist()
            text = df_input[column].astype(str).tolist()
            output_df[column] = [replace_by_dict(t, dictionary, MIN_LEVDIS_SCORE, idx) for idx, t in enumerate(text)]

    output_df.to_excel(output_file, index=False)
