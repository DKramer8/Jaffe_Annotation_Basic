<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0"
    xmlns:tei="http://www.tei-c.org/ns/1.0">

    <xsl:template match="/">
        <xsl:processing-instruction name="xml-model">
            <xsl:text> href="https://raw.githubusercontent.com/hannahbusch/Formierung_Europas_wip/master/tei_cei_formierung_edited.rng" type="application/xml" schematypens="http://relaxng.org/ns/structure/1.0"</xsl:text>
        </xsl:processing-instruction>
        <TEI xmlns="http://www.tei-c.org/ns/1.0">
            <teiHeader>
                <fileDesc>
                    <titleStmt>
                        <title>Converted Document</title>
                    </titleStmt>
                    <publicationStmt>
                        <publisher>Automated Transformation</publisher>
                    </publicationStmt>
                    <sourceDesc>
                        <bibl>Source XML Document</bibl>
                        <diploDesc>
                            <p>
                                <issued>
                                    <date type="issued">
                                        <xsl:value-of select="/data/row/date"/>
                                    </date>
                                    <placeName>
                                        <xsl:value-of select="/data/row/place"/>
                                    </placeName>
                                    <p>
                                        <xsl:value-of select="/data/row/pope"/>
                                    </p>
                                </issued>
                            </p>
                        </diploDesc>
                    </sourceDesc>
                </fileDesc>
            </teiHeader>
            <text>
                <body>
                    <div>
                        <p>
                            <xsl:value-of select="/data/row/number"/>
                        </p>
                        <p>
                            <xsl:value-of select="/data/row/text"/>
                        </p>
                    </div>
                </body>
            </text>
        </TEI>
    </xsl:template>

</xsl:stylesheet>
