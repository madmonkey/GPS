<xsl:stylesheet xmlns:xsl = "http://www.w3.org/1999/XSL/Transform" version = "1.0" >
  <xsl:output method = "xml" standalone ="yes"/>

  <xsl:variable name="deviceId"><![CDATA[/@#{ENTITYDEV};]]></xsl:variable>
  <xsl:variable name="unitId"><![CDATA[/@#{ENTITYUNIT};]]></xsl:variable>

  <!--Copy the existing Xml message except 'CMP' and 'UNT'-->
  <xsl:template match="CMP" />
  <xsl:template match="UNT" />
  <xsl:template match="@*|node()">
    <xsl:copy>
      <xsl:apply-templates select="@*|node()"/>
    </xsl:copy>
  </xsl:template>

  <!--Add or update CMP and UNT fields-->
  <xsl:template match="ID">
    <ID>
      <xsl:apply-templates select="@* | *"/>
      <UNT>
        <xsl:value-of select="$unitId" />
      </UNT>
      <CMP>
        <xsl:value-of select="$deviceId" />
      </CMP>
    </ID>
  </xsl:template>

  <!--Add attribute to rawMessage-->
  <xsl:template match="rawMessage">
    <xsl:copy>
      <xsl:attribute name="method">xml</xsl:attribute>
      <xsl:apply-templates select="@*|node()" />
    </xsl:copy>
  </xsl:template>

  <!--Rebuilding the Message-->
  <xsl:template match="GPSMessage">
    <GPSTransform>
      <MessageStatus>0</MessageStatus>
      <xsl:apply-templates select="@*|node()"/>
      <!--Pastes everything except the GPSMessage element-->
      <!--This gets applied to inside of rawMessage node-->
      <Type>3</Type>
    </GPSTransform>
  </xsl:template>
</xsl:stylesheet>