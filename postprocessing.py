import pandas as pd
import os
import re
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

COLUMNS = ['date', 'pope', 'place']
MIN_LEVDIS_SCORE = 70

#------------------------------------------------------------
#------------------------------------------------------------
# Functions
#------------------------------------------------------------
#------------------------------------------------------------
def replace_by_dict(text: str, dictionary: list, row_nmb: int) -> str:
    '''
    Replaces a given text with the closest match from a provided dictionary if the Levenshtein distance score is above a certain threshold.
    
    :param text: The text to be replaced
    :param dictionary: A list of valid entries to compare against
    :param row_nmb: The row number in the original DataFrame, used for logging purposes
    :return: The original text if no close match is found, or the closest match from the dictionary if the score is above the threshold
    '''

    match, score, _ = process.extractOne(
        text,
        dictionary,
        scorer=fuzz.ratio
    )
    if score >= MIN_LEVDIS_SCORE and score != 100:
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

def correct_number_column(numbers: list) -> list:
    """
    Korrigiert die 'number'-Spalte basierend auf OCR-Output:
    - Die Hauptzahl wird überprüft, ob sie um 1 größer ist als die zuletzt gefundene Hauptzahl.
    - Falls nicht, wird sie korrigiert.
    - Zahlen in Klammern bleiben unverändert.
    
    :param numbers: Liste der originalen 'number'-Werte (Strings)
    :return: Liste der korrigierten 'number'-Werte
    """
    corrected = []
    last_main_number = None  # Keine Zahl bisher

    for n in numbers:
        if pd.isna(n) or str(n).strip() == '':
            corrected.append(n)
            continue

        n_str = str(n).strip()
        # Trenne Hauptzahl und Klammerinhalt
        match = re.match(r'(\d+)?\s*(\(\d+\))?', n_str)
        if match:
            main_num = match.group(1)
            bracket_num = match.group(2) if match.group(2) else ''

            if main_num is not None:
                main_num = int(main_num)
                # Prüfe, ob Hauptzahl korrekt ist
                if last_main_number is not None and main_num != last_main_number + 1:
                    print(f"Korrigiere {main_num} → {last_main_number + 1}")
                    main_num = last_main_number + 1
                last_main_number = main_num
            else:
                main_num = ''  # Falls keine Hauptzahl, leer lassen

            corrected.append(f"{main_num}{bracket_num}")
        else:
            # Unbekanntes Format, Originalwert übernehmen
            corrected.append(n_str)

    return corrected

#------------------------------------------------------------
#------------------------------------------------------------
# Main
#------------------------------------------------------------
#------------------------------------------------------------
if __name__ == '__main__':
    print('Please enter file name (must end in xlsx)')
    user_input = input()
    input_file = INPUT_PATH + user_input
    output_file = OUTPUT_FILE + user_input

    df_dict = pd.read_excel(DICTIONARY_FILE)
    df_input = pd.read_excel(input_file)
    # Nummern funktionieren noch nicht. Bisher aber nur mit ChatGPT probiert, nicht eigenständig
    # if 'number' in df_input.columns:
    #     df_input['number'] = correct_number_column(df_input['number'].tolist())
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

            years = [replace_by_dict(y, year_dict, idx) for idx, y in enumerate(years_raw)]
            months = [replace_by_dict(m, month_dict, idx) for idx, m in enumerate(months_raw)]
            months = [x + '.' if x != ' ' else x for x in months]
            days = [replace_by_dict(d, day_dict, idx) for idx, d in enumerate(days_raw)]
            days = [x + '.' if x != ' ' else x for x in days]

            i = 0
            while i < len(years):
                output_df.loc[i, 'date'] = f'{years[i]} {months[i]} {days[i]}'
                i += 1
        else:
            dictionary = df_dict[column].dropna().astype(str).tolist()
            text = df_input[column].astype(str).tolist()
            output_df[column] = [replace_by_dict(t, dictionary, idx) for idx, t in enumerate(text)]

    output_df.to_excel(output_file, index=False)
