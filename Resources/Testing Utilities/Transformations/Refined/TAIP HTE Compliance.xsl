	<xsl:stylesheet xmlns:xsl = "http://www.w3.org/1999/XSL/Transform" version = "1.0" >
	<!--TAIP HTE Compliance.xsl (c) SunGard HTE Inc--> 
	<!--TAIP HTE Compliance modifies a LAT/LONG transaction (LN, PV) and appends a standard client specific identifier to the message defined in TAIP specification-->
	<!--TAIP HTE Compliance optionally adds a non-standard tag to the message which specifies a status of unit -->
	<!--Added an additional check to check for a valid LAT before passing through-->
	<!--Coordinates should be in Decimal Degrees (DD) since NO conversion is applied here--> 

	<xsl:output method = "xml" standalone ="yes" encoding="UTF-8"/> 
		<xsl:variable name = "transaction" select ="substring(.,2,2)"/>
		<xsl:variable name = "hteid" select = "'/@#{ID};'"/>
		<xsl:variable name = "htestat" select = "'/@#{STATUS};'"/>
		<xsl:template match="/">
			<GPSTransform>
				<xsl:choose>
					<xsl:when test="$transaction='LN'">
						<xsl:choose>
							<xsl:when test ="number(substring(substring-after(.,$transaction),10,9))!=0">
								<MessageStatus>0</MessageStatus>
								<rawMessage>
									<!--spec says 65 is actually 69 + 3 positions for transactions !-->
									<xsl:value-of select="concat(substring(.,1,72),';ID=',$hteid)"/>
									<!--status may be required at a future time
									<xsl:value-of select="concat(';STAT=',$htestat)"/>
									!-->
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
									<xsl:value-of select="concat(substring(.,1,33),';ID=',$hteid)"/>	
									<!--status may be required at a future time
									<xsl:value-of select="concat(';STAT=',$htestat)"/>
									!-->
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