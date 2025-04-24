from lxml import etree
import pandas as pd

INPUT_PATH = "RI_test.xlsx"
OUTPUT_PATH = "RI_Output.xml"
COLUMNS_XML_MAP = { # "excel_column": "xml" 
    "identifier": "idno", # Steht nicht so im Template
    "locality_string": "place_name",
    "start_date": "date_notBefore",
    "end_date": "date_notAfter",
    "summary": "abstract_p",
    "commentary": "commentary_p", 
    "urn": "idno_urn", # Steht nicht so im Template
    "archival_history": "bibliography_list_bibl",
    "recipient": "legal_actor_recipient_inner" # Steht nicht so im Template
}
xml_content_map = { # "xml": "content"
    "title": "[Die Nummer des Regests im Projekt]",
    "resp1": "Ursprünglicher Bearbeiter/Bearbeiter Datenquelle",
    "pers_name1": "[Vorname und Nachname der Bearbeiter*in]",
    "resp2": "Encoding:",
    "pers_name2": "[Vorname und Nachname der Überarbeiter*in/Auszeichnung]",
    "publisher": "Die Formierung Europas",
    "idno": "",
    "idno_urn": "",
    "licence": "Lizenz",
    "bibl": "[Jaffe-Nummer]",
    "place_name": "[Ausstellungsort]",
    "date": "[Ausstellungsdatum]",
    "date_notBefore": "1000-01-01",
    "date_notAfter": "1000-01-01",
    "legal_actor_issuer_inner": "[Issuer]",
    "legal_actor_recipient_inner": "[Recipient 1]",
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
    "bibliography_list_bibl": "[Weitere Referenzen als listBibl]",
    "notes_head": "[Anmerkungen, Notizen und Fragen der Bearbeiter*in.]",
}

def create_tei_xml(output_file, df):
    # Define namespaces
    namespaces = {
        None: "http://www.tei-c.org/ns/1.0",
        "rdf": "http://www.w3.org/1999/02/22-rdf-syntax-ns#"
    }

    # Root element
    tei = etree.Element("TEI", nsmap=namespaces)

    # teiHeader
    tei_header = etree.SubElement(tei, "teiHeader")
    file_desc = etree.SubElement(tei_header, "fileDesc")
    title_stmt = etree.SubElement(file_desc, "titleStmt")
    title = etree.SubElement(title_stmt, "title")
    title.text = xml_content_map["title"]

    resp_stmt1 = etree.SubElement(title_stmt, "respStmt")
    resp1 = etree.SubElement(resp_stmt1, "resp")
    resp1.text = xml_content_map["resp1"]
    pers_name1 = etree.SubElement(resp_stmt1, "persName")
    pers_name1.text = xml_content_map["pers_name1"]

    resp_stmt2 = etree.SubElement(title_stmt, "respStmt")
    resp2 = etree.SubElement(resp_stmt2, "resp", attrib={"when": "2024"})
    resp2.text = xml_content_map["resp2"]
    pers_name2 = etree.SubElement(resp_stmt2, "persName")
    pers_name2.text = xml_content_map["pers_name2"]

    publication_stmt = etree.SubElement(file_desc, "publicationStmt")
    publisher = etree.SubElement(publication_stmt, "publisher")
    publisher.text = xml_content_map["publisher"]
    idno = etree.SubElement(publication_stmt, "idno")
    idno.text = xml_content_map["idno"]
    idno_urn = etree.SubElement(publication_stmt, "idno", attrib={"type": "urn"})
    idno_urn.text = xml_content_map["idno_urn"]    
    availability = etree.SubElement(publication_stmt, "availability")
    licence = etree.SubElement(availability, "licence")
    licence.text = xml_content_map["licence"]

    source_desc = etree.SubElement(file_desc, "sourceDesc")
    bibl = etree.SubElement(source_desc, "bibl", attrib={"type": "jaffe"})
    bibl.text = xml_content_map["bibl"]

    diplo_desc = etree.SubElement(source_desc, "diploDesc")
    issued = etree.SubElement(diplo_desc, "p")
    issued_inner = etree.SubElement(issued, "issued")
    place_name = etree.SubElement(issued_inner, "placeName")
    place_name.text = xml_content_map["place_name"]
    date = etree.SubElement(issued_inner, "date", attrib={
        "type": "issued",
        "notBefore": xml_content_map["date_notBefore"],
        "notAfter": xml_content_map["date_notAfter"]
    })
    date.text = xml_content_map["date"]

    legal_actor_issuer = etree.SubElement(diplo_desc, "p")
    legal_actor_issuer_inner = etree.SubElement(legal_actor_issuer, "legalActor", attrib={"type": "issuer"})
    legal_actor_issuer_inner.text = xml_content_map["legal_actor_issuer_inner"]

    legal_actor_recipient = etree.SubElement(diplo_desc, "p")
    legal_actor_recipient_inner = etree.SubElement(legal_actor_recipient, "legalActor", attrib={"type": "recipient"})
    legal_actor_recipient_inner.text = xml_content_map["legal_actor_recipient_inner"]

    object_type = etree.SubElement(diplo_desc, "p")
    object_type_inner = etree.SubElement(object_type, "objectType")
    object_type_inner.text = xml_content_map["object_type_inner"]

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
    etree.SubElement(abstract_div, "p").text = xml_content_map["abstract_p"]

    # Echtheitskriterien
    authen_div = etree.SubElement(body, "div", attrib={"type": "other"})
    etree.SubElement(authen_div, "head").text = "Echtheitskriterien:"
    authen = etree.SubElement(authen_div, "authen", attrib={"type": "formula"})
    etree.SubElement(authen, "p").text = xml_content_map["authen_p"]

    # Incipit
    incipit_div = etree.SubElement(body, "div")
    etree.SubElement(incipit_div, "head").text = "Incipit:"
    etree.SubElement(incipit_div, "diploPart", attrib={"type": "incipit", "{http://www.w3.org/XML/1998/namespace}lang": "lat"}).text = xml_content_map["incipit_diploPart"]

    # Datatio
    datatio_div = etree.SubElement(body, "div")
    etree.SubElement(datatio_div, "head").text = "Datatio:"
    etree.SubElement(datatio_div, "diploPart", attrib={"type": "datatio"}).text = xml_content_map["datatio_diploPart"]

    # Handschriftliche Überlieferung
    ms_sources_div = etree.SubElement(body, "div", attrib={"type": "msSources"})
    etree.SubElement(ms_sources_div, "head").text = "Handschriftliche Überlieferung und alte Drucke"
    ms_list_bibl = etree.SubElement(ms_sources_div, "listBibl")
    etree.SubElement(ms_list_bibl, "bibl").text = xml_content_map["ms_list_bibl"]

    # Dekretalen
    decretals_div = etree.SubElement(body, "div", attrib={"type": "decretals"})
    etree.SubElement(decretals_div, "head").text = "Dekretalen"
    decretals_list_bibl = etree.SubElement(decretals_div, "listBibl")
    etree.SubElement(decretals_list_bibl, "bibl").text = xml_content_map["decretals_list_bibl"]

    # Historiography
    historiography_div = etree.SubElement(body, "div", attrib={"type": "historiography"})
    etree.SubElement(historiography_div, "head").text = "Erwähnungen in der Historiographie"
    historiography_list_bibl = etree.SubElement(historiography_div, "listBibl")
    etree.SubElement(historiography_list_bibl, "bibl").text = xml_content_map["historiography_list_bibl"]

    # Editions
    editions_div = etree.SubElement(body, "div", attrib={"type": "editions"})
    etree.SubElement(editions_div, "head").text = "Editionen"
    editions_list_bibl = etree.SubElement(editions_div, "listBibl")
    etree.SubElement(editions_list_bibl, "bibl").text = xml_content_map["editions_list_bibl"]

    # Regesta
    regesta_div = etree.SubElement(body, "div", attrib={"type": "regesta"})
    etree.SubElement(regesta_div, "head").text = "Regesten"
    regesta_list_bibl = etree.SubElement(regesta_div, "listBibl")
    etree.SubElement(regesta_list_bibl, "bibl").text = xml_content_map["regesta_list_bibl"]

    # Commentary
    commentary_div = etree.SubElement(body, "div", attrib={"type": "commentary"})
    etree.SubElement(commentary_div, "head").text = "Sachkommentar"
    etree.SubElement(commentary_div, "p").text = xml_content_map["commentary_p"]

    # Bibliography
    bibliography_div = etree.SubElement(body, "div", attrib={"type": "bibliography"})
    etree.SubElement(bibliography_div, "head").text = "Bibliographie"
    bibliography_list_bibl = etree.SubElement(bibliography_div, "listBibl")
    etree.SubElement(bibliography_list_bibl, "bibl").text = xml_content_map["bibliography_list_bibl"]

    # Notes
    notes_div = etree.SubElement(body, "div", attrib={"type": "note"})
    etree.SubElement(notes_div, "head").text = xml_content_map["notes_head"]

    # Write to file with processing instructions
    with open(output_file, "wb") as f:
        f.write(b'<?xml version="1.0" encoding="UTF-8"?>\n')
        f.write(b'<?xml-model href="https://raw.githubusercontent.com/cceh/FormierungEuropas/refs/heads/main/CEI_TEI_Schema/tei_cei_formierung_edited.rng" type="application/xml" schematypens="http://relaxng.org/ns/structure/1.0"?>\n')
        f.write(b'<?xml-stylesheet type="text/css" href="https://raw.githubusercontent.com/hannahbusch/Formierung_Europas_wip/master/style.css"?>\n')
        f.write(etree.tostring(tei, pretty_print=True, xml_declaration=False, encoding="UTF-8"))

df = pd.read_excel(INPUT_PATH, header=0, dtype=str)
for id, row in df.iterrows():
    for col, value in row.items():
        if col in COLUMNS_XML_MAP and value != "nan":
            try:
                xml_content_map[COLUMNS_XML_MAP[col]] = value
            except:
                print(f"Spalte '{col}' nicht in xml_content_map gefunden.")
create_tei_xml(OUTPUT_PATH, df)
print(f"XML-Datei wurde erstellt: {OUTPUT_PATH}")