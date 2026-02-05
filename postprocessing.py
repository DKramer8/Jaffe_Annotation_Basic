#Postprocessing der Päpste, Jahre?, Monate, Orte, Nummern? Und leere Spalten hinzufügen

# Testing....
from rapidfuzz import process, fuzz

dictionary = {
    "apfel": 1,
    "banane": 2,
    "orange": 3,
    "birne": 4
}

def ersetze_durch_aehnlichstes_wort(text, dictionary, min_score=70):
    match, score, _ = process.extractOne(
        text,
        dictionary.keys(),
        scorer=fuzz.ratio
    )
    return match if score >= min_score else text

eingabe = "apfl"
print(ersetze_durch_aehnlichstes_wort(eingabe, dictionary))
