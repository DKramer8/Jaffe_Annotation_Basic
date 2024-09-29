<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0"
    xmlns:tei="http://www.tei-c.org/ns/1.0">

    <xsl:template match="/">
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
                    </sourceDesc>
                </fileDesc>
            </teiHeader>
            <text>
                <body>
                    <xsl:for-each select="data/row">
                        <div n="{index}">
                            <sourceDesc>
                                <diploDesc>
                                    <p>
                                        <issued>
                                            <date type="issued" notBefore="" notAfter="">
                                                <xsl:value-of select="date"/>
                                            </date>
                                            <placeName>
                                                <xsl:value-of select="place"/>
                                            </placeName>
                                        </issued>
                                    </p>
                                </diploDesc>
                            </sourceDesc>
                            <number type="JL">
                                <xsl:value-of select="number"/>
                            </number>
                            <div type="abstract">
                                <p>
                                    <xsl:value-of select="text"/>
                                </p>
                            </div>
                        </div>
                    </xsl:for-each>
                </body>
            </text>
        </TEI>
    </xsl:template>

</xsl:stylesheet>
