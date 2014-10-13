	<xsl:stylesheet xmlns:xsl = "http://www.w3.org/1999/XSL/Transform" version = "1.0" >
	<!--No Remote NMEA.xsl (c) SunGard HTE Inc--> 
	<!--Discards all messages regardless of content, used to short-circuit remote clients not identifying themselves-->
	<xsl:output method = "xml" standalone ="yes" encoding="UTF-8"/> 
		<xsl:template match="/">
			<GPSTransform>
				<MessageStatus>3</MessageStatus>
				<rawMessage>NOT A VALID PROTOCOL</rawMessage>
				<Type>1</Type>
			</GPSTransform>
		</xsl:template>
	</xsl:stylesheet>