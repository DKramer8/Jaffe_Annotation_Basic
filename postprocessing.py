#Postprocessing der Jahre, Monate und Nummern noch hinzufügen. Und leere Spalten hinzufügen
import pandas as pd
import os
from rapidfuzz import process, fuzz

DIR_PATH = os.getcwd()
DICTIONARY_FILE = DIR_PATH + f'/data/static/jaffe_dicts.xlsx'
INPUT_FILE = DIR_PATH + f'/data/output/jaffe2_preAlexander.xlsx'
OUTPUT_FILE = f'postprocessed.xlsx'
COLUMNS = ['pope', 'place']

def ersetze_durch_aehnlichstes_wort(text, dictionary, row_nmb, min_score=70):
    match, score, _ = process.extractOne(
        text,
        dictionary,
        scorer=fuzz.ratio
    )
    if score >= min_score and score != 100:
        print(f"Ersetze '{text}' durch '{match}' (Score: {score}) at row {row_nmb+2}")
        return match
    else:
        return text

if __name__ == '__main__':
    df_dict = pd.read_excel(DICTIONARY_FILE)
    df_input = pd.read_excel(INPUT_FILE)
    output_df = df_input

    for column in COLUMNS:
        text = df_input[column].astype(str).tolist()
        dictionary = df_dict[column].dropna().astype(str).tolist()
        output_df[column] = [ersetze_durch_aehnlichstes_wort(t, dictionary, idx) for idx, t in enumerate(text)]

    output_df.to_excel(OUTPUT_FILE, index=False)
