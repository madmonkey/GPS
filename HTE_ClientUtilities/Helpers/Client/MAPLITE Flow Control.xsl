	<xsl:stylesheet xmlns:xsl = "http://www.w3.org/1999/XSL/Transform" version = "1.0" >
	<!--MAPLITE Flow Control.xsl (c) SunGard HTE Inc--> 
	<!--MAPLITE Flow Control.xsl reflects on when a message was last processed before allowing it to be transmitted-->
	<xsl:output method = "xml" standalone ="yes" encoding="UTF-8"/> 
		<xsl:variable name="lastProcessed" select="/@#{TIMEELAPSED};"/>
		<xsl:template match="/">
			<GPSTransform>
				<xsl:choose>
					<xsl:when test="number($lastProcessed) > 30">
						<MessageStatus>0</MessageStatus>
						<rawMessage>
							<xsl:value-of select="/"/>
						</rawMessage>
						<Type>2</Type>
					</xsl:when>
					<xsl:otherwise>
						<MessageStatus>3</MessageStatus>
						<rawMessage>MESSAGE DROPPED FOR FLOW CONTROL</rawMessage>
						<Type>2</Type>
					</xsl:otherwise>
				</xsl:choose>
			</GPSTransform>
		</xsl:template>
	</xsl:stylesheet>