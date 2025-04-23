from lxml import etree

EXCEL_XML_MAP = { # "ColumnName": "XMLTagName"
    "identifier": "idno",
    "locality_string": "placeName",
    "start_date": "notBefore",
    "end_date": "notAfter",
    "summary": "p",
    "commentary": "p",
    "urn": "idno"
}

def create_tei_xml(output_file):
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
    title.text = "[Die Nummer des Regests im Projekt]"

    resp_stmt1 = etree.SubElement(title_stmt, "respStmt")
    resp1 = etree.SubElement(resp_stmt1, "resp")
    resp1.text = "Ursprünglicher Bearbeiter/Bearbeiter Datenquelle"
    pers_name1 = etree.SubElement(resp_stmt1, "persName")
    pers_name1.text = "[Vorname und Nachname der Bearbeiter*in]"

    resp_stmt2 = etree.SubElement(title_stmt, "respStmt")
    resp2 = etree.SubElement(resp_stmt2, "resp", attrib={"when": "2024"})
    resp2.text = "Encoding:"
    pers_name2 = etree.SubElement(resp_stmt2, "persName")
    pers_name2.text = "[Vorname und Nachname der Überarbeiter*in/Auszeichnung]"

    publication_stmt = etree.SubElement(file_desc, "publicationStmt")
    publisher = etree.SubElement(publication_stmt, "publisher")
    publisher.text = "Die Formierung Europas"
    availability = etree.SubElement(publication_stmt, "availability")
    licence = etree.SubElement(availability, "licence")
    licence.text = "Lizenz"

    source_desc = etree.SubElement(file_desc, "sourceDesc")
    bibl = etree.SubElement(source_desc, "bibl", attrib={"type": "jaffe"})
    bibl.text = "[Jaffe-Nummer]"

    diplo_desc = etree.SubElement(source_desc, "diploDesc")
    issued = etree.SubElement(diplo_desc, "p")
    issued_inner = etree.SubElement(issued, "issued")
    place_name = etree.SubElement(issued_inner, "placeName")
    place_name.text = "[Ausstellungsort]"
    date = etree.SubElement(issued_inner, "date", attrib={"type": "issued"})
    date.text = "[Ausstellungsdatum]"

    legal_actor_issuer = etree.SubElement(diplo_desc, "p")
    legal_actor_issuer_inner = etree.SubElement(legal_actor_issuer, "legalActor", attrib={"type": "issuer"})
    legal_actor_issuer_inner.text = "[Issuer]"

    legal_actor_recipient = etree.SubElement(diplo_desc, "p")
    legal_actor_recipient_inner = etree.SubElement(legal_actor_recipient, "legalActor", attrib={"type": "recipient"})
    legal_actor_recipient_inner.text = "[Recipient 1]"

    object_type = etree.SubElement(diplo_desc, "p")
    object_type_inner = etree.SubElement(object_type, "objectType")
    object_type_inner.text = "[Typisierung 1]"

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
    etree.SubElement(abstract_div, "p").text = "[Regestentext kann diploParts mit Sprachangabe enthalten.]"

    # Echtheitskriterien
    authen_div = etree.SubElement(body, "div", attrib={"type": "other"})
    etree.SubElement(authen_div, "head").text = "Echtheitskriterien:"
    authen = etree.SubElement(authen_div, "authen", attrib={"type": "formula"})
    etree.SubElement(authen, "p").text = "[*+?-]"

    # Incipit
    incipit_div = etree.SubElement(body, "div")
    etree.SubElement(incipit_div, "head").text = "Incipit:"
    etree.SubElement(incipit_div, "diploPart", attrib={"type": "incipit", "{http://www.w3.org/XML/1998/namespace}lang": "lat"}).text = "[Lorem ipsum dolor sit amet]"

    # Datatio
    datatio_div = etree.SubElement(body, "div")
    etree.SubElement(datatio_div, "head").text = "Datatio:"
    etree.SubElement(datatio_div, "diploPart", attrib={"type": "datatio"}).text = "-"

    # Handschriftliche Überlieferung
    ms_sources_div = etree.SubElement(body, "div", attrib={"type": "msSources"})
    etree.SubElement(ms_sources_div, "head").text = "Handschriftliche Überlieferung und alte Drucke"
    ms_list_bibl = etree.SubElement(ms_sources_div, "listBibl")
    etree.SubElement(ms_list_bibl, "bibl").text = "[Handschrift oder Alter Druck]"
    etree.SubElement(ms_list_bibl, "bibl").text = "[Handschrift oder Alter Druck]"

    # Dekretalen
    decretals_div = etree.SubElement(body, "div", attrib={"type": "decretals"})
    etree.SubElement(decretals_div, "head").text = "Dekretalen"
    decretals_list_bibl = etree.SubElement(decretals_div, "listBibl")
    etree.SubElement(decretals_list_bibl, "bibl").text = "[Dekretale 1]"
    etree.SubElement(decretals_list_bibl, "bibl").text = "[Dekretale 2]"

    # Historiography
    historiography_div = etree.SubElement(body, "div", attrib={"type": "historiography"})
    etree.SubElement(historiography_div, "head").text = "Erwähnungen in der Historiographie"
    historiography_list_bibl = etree.SubElement(historiography_div, "listBibl")
    etree.SubElement(historiography_list_bibl, "bibl").text = "[i.d.R. nach Edition, also Literaturtitel zitiert]"

    # Editions
    editions_div = etree.SubElement(body, "div", attrib={"type": "editions"})
    etree.SubElement(editions_div, "head").text = "Editionen"
    editions_list_bibl = etree.SubElement(editions_div, "listBibl")
    etree.SubElement(editions_list_bibl, "bibl").text = "[Edition 1]"
    etree.SubElement(editions_list_bibl, "bibl").text = "[Edition 2]"
    etree.SubElement(editions_list_bibl, "bibl").text = "[Edition 3]"

    # Regesta
    regesta_div = etree.SubElement(body, "div", attrib={"type": "regesta"})
    etree.SubElement(regesta_div, "head").text = "Regesten"
    regesta_list_bibl = etree.SubElement(regesta_div, "listBibl")
    etree.SubElement(regesta_list_bibl, "bibl").text = "[Regest 1]"
    etree.SubElement(regesta_list_bibl, "bibl").text = "[Regest 2]"
    etree.SubElement(regesta_list_bibl, "bibl").text = "[Regest 3]"

    # Commentary
    commentary_div = etree.SubElement(body, "div", attrib={"type": "commentary"})
    etree.SubElement(commentary_div, "head").text = "Sachkommentar"
    etree.SubElement(commentary_div, "p").text = "[Inhalt des Sachkommentars. Bibliographische Referenzen werden als <bibl>Bibliographische Angabe</bibl> getagged.]"

    # Bibliography
    bibliography_div = etree.SubElement(body, "div", attrib={"type": "bibliography"})
    etree.SubElement(bibliography_div, "head").text = "Bibliographie"
    bibliography_list_bibl = etree.SubElement(bibliography_div, "listBibl")
    etree.SubElement(bibliography_list_bibl, "bibl").text = "[Weitere Referenzen als listBibl]"

    # Notes
    notes_div = etree.SubElement(body, "div", attrib={"type": "note"})
    etree.SubElement(notes_div, "head").text = "[Anmerkungen, Notizen und Fragen der Bearbeiter*in.]"

    # Write to file with processing instructions
    with open(output_file, "wb") as f:
        f.write(b'<?xml version="1.0" encoding="UTF-8"?>\n')
        f.write(b'<?xml-model href="https://raw.githubusercontent.com/cceh/FormierungEuropas/refs/heads/main/CEI_TEI_Schema/tei_cei_formierung_edited.rng" type="application/xml" schematypens="http://relaxng.org/ns/structure/1.0"?>\n')
        f.write(b'<?xml-stylesheet type="text/css" href="https://raw.githubusercontent.com/hannahbusch/Formierung_Europas_wip/master/style.css"?>\n')
        f.write(etree.tostring(tei, pretty_print=True, xml_declaration=False, encoding="UTF-8"))

# Usage
output_path = "output.xml"
create_tei_xml(output_path)
print(f"XML-Datei wurde erstellt: {output_path}")