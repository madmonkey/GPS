<xsl:stylesheet xmlns:xsl = "http://www.w3.org/1999/XSL/Transform" version = "1.0" > 
				<xsl:output method = "xml" standalone ="yes" encoding="UTF-8"/> 
					<xsl:variable name = "transaction" select ="substring(.,2,2)"/>
					<xsl:template match="/">
					<GPSTransform>
						<xsl:choose>
						<xsl:when test="$transaction='LN'">
							<MessageStatus>0</MessageStatus>
							<rawMessage>
								<xsl:value-of select="/"/>
							</rawMessage>
							<Type>0</Type>
						</xsl:when>
						<xsl:when test="$transaction='PV'">
							<MessageStatus>0</MessageStatus>
							<rawMessage>
								<xsl:value-of select="/"/>
							</rawMessage>
							<Type>0</Type>
						</xsl:when>
						<xsl:otherwise>
							<MessageStatus>3</MessageStatus>
							<rawMessage>NOT A VALID TRANSACTION TYPE</rawMessage>
							<Type>0</Type>
						</xsl:otherwise>
						</xsl:choose>
					</GPSTransform>
					</xsl:template>
				</xsl:stylesheet>