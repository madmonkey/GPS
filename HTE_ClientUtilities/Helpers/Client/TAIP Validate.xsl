	<xsl:stylesheet xmlns:xsl = "http://www.w3.org/1999/XSL/Transform" version = "1.0" >
	<!--TAIP Validate.xsl (c) SunGard HTE Inc--> 
	<!--TAIP Validate checks for valid LAT/LONG transaction (LN, PV) ensures that no other transaction goes to CAD-->
	<!--Added an additional check to check for a valid LAT before passing through-->
	<!--Coordinates must be in Decimal Degrees (DD) since NO conversion is applied here--> 
	<xsl:output method = "xml" standalone ="yes" encoding="UTF-8"/> 
		<xsl:variable name = "transaction" select ="substring(.,2,2)"/>
		<xsl:template match="/">
			<GPSTransform>
				<xsl:choose>
					<xsl:when test="$transaction='LN'">
						<xsl:choose>
							<xsl:when test ="number(substring(substring-after(.,$transaction),10,9))!=0">
								<MessageStatus>0</MessageStatus>
								<rawMessage>
									<xsl:value-of select="/"/>
								</rawMessage>
								<Type>0</Type>
							</xsl:when>
							<xsl:otherwise>
								<MessageStatus>3</MessageStatus>
								<rawMessage>NO VALID DATA</rawMessage>
								<Type>0</Type>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:when test="$transaction='PV'">
						<xsl:choose>
							<xsl:when test ="number(substring(substring-after(.,$transaction),7,6))!=0">
								<MessageStatus>0</MessageStatus>
								<rawMessage>
									<xsl:value-of select="/"/>
								</rawMessage>
								<Type>0</Type>
							</xsl:when>
							<xsl:otherwise>
								<MessageStatus>3</MessageStatus>
								<rawMessage>NO VALID DATA</rawMessage>
								<Type>0</Type>
							</xsl:otherwise>
						</xsl:choose>
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