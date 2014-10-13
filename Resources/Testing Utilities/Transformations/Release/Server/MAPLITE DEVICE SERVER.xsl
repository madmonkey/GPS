	<xsl:stylesheet xmlns:xsl = "http://www.w3.org/1999/XSL/Transform" version = "1.0" >
	<!--MAPLITE DEVICE SERVER.xsl (c) SunGard HTE Inc--> 
	<!--Reconciles Smart-client device identification with server settings-->
	<xsl:output method = "xml" standalone ="yes" encoding="UTF-8"/> 
		<xsl:template match="/">
			<GPSTransform>
				<xsl:variable name="unitID"><![CDATA[/@#{ENTITYDEV};]]></xsl:variable>
				<MessageStatus>0</MessageStatus>
				<rawMessage>
					<xsl:value-of select ="substring(.,1,20)"/>
					<xsl:value-of select ="substring(concat($unitID,'                '),1,16)"/>
					<xsl:value-of select ="substring(.,37)"/>
				</rawMessage>
				<Type>2</Type>
			</GPSTransform>
		</xsl:template>
	</xsl:stylesheet>