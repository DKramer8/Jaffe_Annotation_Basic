from lxml import etree
import pandas as pd
from bs4 import BeautifulSoup

INPUT_PATH = "input/Regesta-Imperii_Papsturkunden_sondiert_IB_JM_JB_SG_import.xlsx"
OUTPUT_FOLDER = "output"
REGEST_TYPE = "papst"
COLUMNS_XML_MAP = { # "excel_column": "xml" 
    "identifier": "idno",
    "place": "place_name",
    "notBefore": "date_notBefore",
    "notAfter": "date_notAfter",
    "abstract": "abstract_p",
    "bibliography": "archival_history_bibl",# -> archival_history
    "sourceDesc": "bibl", # Plus den identifier; ist so in der Logik unten eingbaut
    "commentary": "commentary_p",
    "literature": "literature_bibl", # neuer div wie bei bibliografie
    "footnotes": "notes_p", # footnotes -> div notes
    "annotations": "note", # -> <note><p>
    "incipit": "incipit_diploPart",
    "original_date": "datatio_diploPart", #  -> datatio
    "seal": "seal_p", # -> div note
    "recipient": "legal_actor_recipient_inner",
    "witnesses": "legal_actor_witness_inner", # -> zu anderen legalActors nur als witness
    "clerk": "legal_actor_clerk_inner", # -> wie oben
    "chancellor": "legal_actor_chancellor_inner", # -> wie oben
    "external_links": "exLinks_bibl", # div bibliografie <bibl>
    # "exchange_identifier": "", # Ist zumindest in de Excels gefüllt, aber keine Ahnung wie wichtig das ist
    "urn": "idno_urn", 
    "date_string": "date",
}
XML_CONTENT_MAP = { # "xml": "content"
    "title": "[Die Nummer des Regests im Projekt]",
    "resp1": "Ursprünglicher Bearbeiter/Bearbeiter Datenquelle",
    "pers_name1": "[Vorname und Nachname der Bearbeiter*in]",
    "resp2": "Encoding:",
    "pers_name2": "[Vorname und Nachname der Überarbeiter*in/Auszeichnung]",
    "publisher": "Die Formierung Europas",
    "idno": "",
    "idno_urn": "",
    "licence": "Lizenz",
    "bibl": "",
    "bibl_jaffe": "[Jaffe-Nummer]",
    "place_name": "[Ausstellungsort]",
    "date": "[Ausstellungsdatum]",
    "date_notBefore": "1000-01-01",
    "date_notAfter": "1000-01-01",
    "legal_actor_issuer_inner": "[Issuer]",
    "legal_actor_recipient_inner": "[Recipient 1]",
    "legal_actor_witness_inner": "",
    "legal_actor_clerk_inner": "",
    "legal_actor_chancellor_inner": "",
    "object_type_inner": "[Typisierung 1]",
    "abstract_p": "[Regestentext kann diploParts mit Sprachangabe enthalten.]",
    "authen_p": "[*+?-]",
    "incipit_diploPart": "[Lorem ipsum dolor sit amet]",
    "datatio_diploPart": "-",
    "ms_list_bibl": "[Handschrift oder Alter Druck]",
    "decretals_list_bibl": "[Dekretale 1]",
    "historiography_list_bibl": "[i.d.R. nach Edition, also Literaturtitel zitiert]",
    "editions_list_bibl": "[Edition 1]",
    "regesta_list_bibl": "[Regest 1]",
    "commentary_p": "[Inhalt des Sachkommentars. Bibliographische Referenzen werden als <bibl>Bibliographische Angabe</bibl> getagged.]",
    "archival_history_bibl": "[Weitere Referenzen als listBibl]",
    "literature_bibl": "[Literaturverzeichnis]",
    "exLinks_bibl": "[Externe Links]",
    "notes_p": "[Anmerkungen, Notizen und Fragen der Bearbeiter*in.]",
    "seal_p": "",
    "note": "",
}

def remove_html_tags(text):
    if isinstance(text, str):
        return BeautifulSoup(text, "html.parser").get_text()
    return text

def create_tei_xml(output_file):
    # Define namespaces
    namespaces = {
        None: "http://www.tei-c.org/ns/1.0",
        "rdf": "http://www.w3.org/1999/02/22-rdf-syntax-ns#"
    }

    # Root element
    tei = etree.Element("TEI", nsmap=namespaces, source="formierung-europas_regesta_imperii")

    # teiHeader
    tei_header = etree.SubElement(tei, "teiHeader")
    file_desc = etree.SubElement(tei_header, "fileDesc")
    title_stmt = etree.SubElement(file_desc, "titleStmt")
    title = etree.SubElement(title_stmt, "title")
    title.text = str(xml_content_map_instance["title"])

    resp_stmt1 = etree.SubElement(title_stmt, "respStmt")
    resp1 = etree.SubElement(resp_stmt1, "resp")
    resp1.text = str(xml_content_map_instance["resp1"])
    pers_name1 = etree.SubElement(resp_stmt1, "persName")
    pers_name1.text = str(xml_content_map_instance["pers_name1"])

    resp_stmt2 = etree.SubElement(title_stmt, "respStmt")
    resp2 = etree.SubElement(resp_stmt2, "resp", attrib={"when": "2024"})
    resp2.text = str(xml_content_map_instance["resp2"])
    pers_name2 = etree.SubElement(resp_stmt2, "persName")
    pers_name2.text = str(xml_content_map_instance["pers_name2"])

    publication_stmt = etree.SubElement(file_desc, "publicationStmt")
    publisher = etree.SubElement(publication_stmt, "publisher")
    publisher.text = str(xml_content_map_instance["publisher"])
    idno = etree.SubElement(publication_stmt, "idno")
    idno.text = str(xml_content_map_instance["idno"])
    idno_urn = etree.SubElement(publication_stmt, "idno", attrib={"type": "urn"})
    idno_urn.text = str(xml_content_map_instance["idno_urn"])
    availability = etree.SubElement(publication_stmt, "availability")
    licence = etree.SubElement(availability, "licence")
    licence.text = str(xml_content_map_instance["licence"])

    source_desc = etree.SubElement(file_desc, "sourceDesc")
    bibl = etree.SubElement(source_desc, "bibl")
    bibl.text = f"{str(xml_content_map_instance['bibl'])}, {str(xml_content_map_instance['idno'])}"
    bibl_jaffe = etree.SubElement(source_desc, "bibl", attrib={"type": "jaffe"})
    bibl_jaffe.text = str(xml_content_map_instance["bibl_jaffe"])

    diplo_desc = etree.SubElement(source_desc, "diploDesc")
    issued = etree.SubElement(diplo_desc, "p")
    issued_inner = etree.SubElement(issued, "issued")
    place_name = etree.SubElement(issued_inner, "placeName")
    place_name.text = str(xml_content_map_instance["place_name"])
    date = etree.SubElement(issued_inner, "date", attrib={
        "type": "issued",
        "notBefore": str(xml_content_map_instance["date_notBefore"]),
        "notAfter": str(xml_content_map_instance["date_notAfter"])
    })
    date.text = str(xml_content_map_instance["date"])

    legal_actor_issuer = etree.SubElement(diplo_desc, "p")
    legal_actor_issuer_inner = etree.SubElement(legal_actor_issuer, "legalActor", attrib={"type": "issuer"})
    legal_actor_issuer_inner.text = str(xml_content_map_instance["legal_actor_issuer_inner"])

    legal_actor_recipient = etree.SubElement(diplo_desc, "p")
    legal_actor_recipient_inner = etree.SubElement(legal_actor_recipient, "legalActor", attrib={"type": "recipient"})
    legal_actor_recipient_inner.text = str(xml_content_map_instance["legal_actor_recipient_inner"])

    legal_actor_witness = etree.SubElement(diplo_desc, "p")
    legal_actor_witness_inner = etree.SubElement(legal_actor_witness, "legalActor", attrib={"type": "witness"})
    legal_actor_witness_inner.text = str(xml_content_map_instance["legal_actor_witness_inner"])

    legal_actor_clerk = etree.SubElement(diplo_desc, "p")
    legal_actor_clerk_inner = etree.SubElement(legal_actor_clerk, "legalActor")
    legal_actor_clerk_inner.text = str(xml_content_map_instance["legal_actor_clerk_inner"]) 

    legal_actor_chancellor = etree.SubElement(diplo_desc, "p")
    legal_actor_chancellor_inner = etree.SubElement(legal_actor_chancellor, "legalActor")
    legal_actor_chancellor_inner.text = str(xml_content_map_instance["legal_actor_chancellor_inner"])        

    object_type = etree.SubElement(diplo_desc, "p")
    object_type_inner = etree.SubElement(object_type, "objectType")
    object_type_inner.text = str(xml_content_map_instance["object_type_inner"])

    # profileDesc
    profile_desc = etree.SubElement(tei_header, "profileDesc")
    lang_usage = etree.SubElement(profile_desc, "langUsage")
    etree.SubElement(lang_usage, "language", attrib={"ident": "ger"}).text = "German"
    etree.SubElement(lang_usage, "language", attrib={"ident": "lat"}).text = "Latin"
    etree.SubElement(lang_usage, "language", attrib={"ident": "en-GB"}).text = "English"

    # xenoData
    xeno_data = etree.SubElement(tei_header, "xenoData")
    rdf = etree.SubElement(xeno_data, "{http://www.w3.org/1999/02/22-rdf-syntax-ns#}RDF", nsmap={
        "rdf": "http://www.w3.org/1999/02/22-rdf-syntax-ns#",
        "rdfs": "http://www.w3.org/2000/01/rdf-schema#",
        "as": "http://www.w3.org/ns/activitystreams#",
        "cwrc": "http://sparql.cwrc.ca/ontologies/cwrc#",
        "dc": "http://purl.org/dc/elements/1.1/",
        "dcterms": "http://purl.org/dc/terms/",
        "foaf": "http://xmlns.com/foaf/0.1/",
        "geo": "http://www.geonames.org/ontology#",
        "oa": "http://www.w3.org/ns/oa#",
        "schema": "http://schema.org/",
        "xsd": "http://www.w3.org/2001/XMLSchema#",
        "fabio": "https://purl.org/spar/fabio#",
        "bf": "http://www.openlinksw.com/schemas/bif#",
        "cito": "https://sparontologies.github.io/cito/current/cito.html#",
        "org": "http://www.w3.org/ns/org#"
    })

    # text
    text = etree.SubElement(tei, "text")
    body = etree.SubElement(text, "body")

    # Abstract
    abstract_div = etree.SubElement(body, "div", attrib={"type": "abstract"})
    etree.SubElement(abstract_div, "head").text = "Abstract:"
    etree.SubElement(abstract_div, "p").text = str(xml_content_map_instance["abstract_p"])

    # Echtheitskriterien
    authen_div = etree.SubElement(body, "div", attrib={"type": "other"})
    etree.SubElement(authen_div, "head").text = "Echtheitskriterien:"
    authen = etree.SubElement(authen_div, "authen", attrib={"type": "formula"})
    etree.SubElement(authen, "p").text = str(xml_content_map_instance["authen_p"])

    # Incipit
    incipit_div = etree.SubElement(body, "div")
    etree.SubElement(incipit_div, "head").text = "Incipit:"
    etree.SubElement(incipit_div, "diploPart", attrib={"type": "incipit", "{http://www.w3.org/XML/1998/namespace}lang": "lat"}).text = str(xml_content_map_instance["incipit_diploPart"])

    # Datatio
    datatio_div = etree.SubElement(body, "div")
    etree.SubElement(datatio_div, "head").text = "Datatio:"
    etree.SubElement(datatio_div, "diploPart", attrib={"type": "datatio"}).text = str(xml_content_map_instance["datatio_diploPart"])

    # Handschriftliche Überlieferung
    ms_sources_div = etree.SubElement(body, "div", attrib={"type": "msSources"})
    etree.SubElement(ms_sources_div, "head").text = "Handschriftliche Überlieferung und alte Drucke"
    ms_list_bibl = etree.SubElement(ms_sources_div, "listBibl")
    etree.SubElement(ms_list_bibl, "bibl").text = str(xml_content_map_instance["ms_list_bibl"])

    # Dekretalen
    decretals_div = etree.SubElement(body, "div", attrib={"type": "decretals"})
    etree.SubElement(decretals_div, "head").text = "Dekretalen"
    decretals_list_bibl = etree.SubElement(decretals_div, "listBibl")
    etree.SubElement(decretals_list_bibl, "bibl").text = str(xml_content_map_instance["decretals_list_bibl"])

    # Historiography
    historiography_div = etree.SubElement(body, "div", attrib={"type": "historiography"})
    etree.SubElement(historiography_div, "head").text = "Erwähnungen in der Historiographie"
    historiography_list_bibl = etree.SubElement(historiography_div, "listBibl")
    etree.SubElement(historiography_list_bibl, "bibl").text = str(xml_content_map_instance["historiography_list_bibl"])

    # Editions
    editions_div = etree.SubElement(body, "div", attrib={"type": "editions"})
    etree.SubElement(editions_div, "head").text = "Editionen"
    editions_list_bibl = etree.SubElement(editions_div, "listBibl")
    etree.SubElement(editions_list_bibl, "bibl").text = str(xml_content_map_instance["editions_list_bibl"])

    # Regesta
    regesta_div = etree.SubElement(body, "div", attrib={"type": "regesta"})
    etree.SubElement(regesta_div, "head").text = "Regesten"
    regesta_list_bibl = etree.SubElement(regesta_div, "listBibl")
    etree.SubElement(regesta_list_bibl, "bibl").text = str(xml_content_map_instance["regesta_list_bibl"])

    # Commentary
    commentary_div = etree.SubElement(body, "div", attrib={"type": "commentary"})
    etree.SubElement(commentary_div, "head").text = "Sachkommentar"
    etree.SubElement(commentary_div, "p").text = str(xml_content_map_instance["commentary_p"])

    # Bibliography
    archival_history_div = etree.SubElement(body, "div", attrib={"type": "bibliography"})
    etree.SubElement(archival_history_div, "head").text = "Archival History"
    archival_history_bibl = etree.SubElement(archival_history_div, "listBibl")
    etree.SubElement(archival_history_bibl, "bibl").text = str(xml_content_map_instance["archival_history_bibl"])

    # Literature
    literature_div = etree.SubElement(body, "div", attrib={"type": "bibliography"})
    etree.SubElement(literature_div, "head").text = "Literatur"
    literature_bibl = etree.SubElement(literature_div, "listBibl")
    etree.SubElement(literature_bibl, "bibl").text = str(xml_content_map_instance["literature_bibl"])    

    # External Links
    exLinks_div = etree.SubElement(body, "div", attrib={"type": "bibliography"})
    etree.SubElement(exLinks_div, "head").text = "Externe Links"
    exLinks_bibl = etree.SubElement(exLinks_div, "listBibl")
    etree.SubElement(exLinks_bibl, "bibl").text = str(xml_content_map_instance["exLinks_bibl"])      

    # Notes
    notes_div = etree.SubElement(body, "div", attrib={"type": "note"})
    etree.SubElement(notes_div, "head").text = "Fußnoten"
    etree.SubElement(notes_div, "p").text = str(xml_content_map_instance["notes_p"])

    # Seal
    seal_div = etree.SubElement(body, "div", attrib={"type": "note"})
    etree.SubElement(seal_div, "p").text = str(xml_content_map_instance["seal_p"])

    # Single Note Element
    note = etree.SubElement(body, "note")
    etree.SubElement(note, "p").text = str(xml_content_map_instance["note"])

    # Write to file with processing instructions
    with open(output_file, "wb") as f:
        f.write(b'<?xml version="1.0" encoding="UTF-8"?>\n')
        f.write(b'<?xml-model href="https://raw.githubusercontent.com/cceh/FormierungEuropas/refs/heads/main/CEI_TEI_Schema/tei_cei_formierung_edited.rng" type="application/xml" schematypens="http://relaxng.org/ns/structure/1.0"?>\n')
        f.write(b'<?xml-stylesheet type="text/css" href="https://raw.githubusercontent.com/hannahbusch/Formierung_Europas_wip/master/style.css"?>\n')
        f.write(etree.tostring(tei, pretty_print=True, xml_declaration=False, encoding="UTF-8"))

df = pd.read_excel(INPUT_PATH, header=0, dtype=str)
for id, row in df.iterrows():
    xml_content_map_instance = XML_CONTENT_MAP.copy()
    for col, value in row.items():
        if col in COLUMNS_XML_MAP and str(value) != "nan" and value is not None and not pd.isna(value):
            clean_value = remove_html_tags(value)
            try:
                xml_content_map_instance[COLUMNS_XML_MAP[col]] = clean_value
            except:
                print(f"{id}: Spalte '{col}' nicht in xml_content_map_instance gefunden.")
    output_file = f"{OUTPUT_FOLDER}/RI_{REGEST_TYPE}_{id}.xml"
    create_tei_xml(output_file)
    print(f"XML-Datei wurde erstellt: {output_file}")

#output_file = f"{OUTPUT_FOLDER}/RI_{REGEST_TYPE}_{id}.xml"
#create_tei_xml(output_file)
#print(f"XML-Datei wurde erstellt: {output_file}")