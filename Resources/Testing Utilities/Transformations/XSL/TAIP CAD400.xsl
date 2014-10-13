<xsl:stylesheet xmlns:xsl = "http://www.w3.org/1999/XSL/Transform" version = "1.0" > 
				<xsl:output method = "xml" standalone ="yes" encoding="UTF-8"/> 
					<!--When transmitting coordinates to the CAD400 it expects Decimal coordinates if the device transmits as Sexagesimal use this transformation!-->
					<!--CAD400 only accepts PV transactions, if one is so inclined you may add a conversion function here to transform from LN (LEFT AS EXERCISE TO READER) !-->
					<xsl:variable name = "transaction" select ="substring(.,2,2)"/>
					<xsl:variable name = "toDecimal" select = "'true'"/>
					<xsl:template match="/">
					<GPSTransform>
						<xsl:choose>
						<xsl:when test="$transaction='PV'">
							<xsl:choose>
								<xsl:when test="$toDecimal='true'">
									<!--CONVERT FROM STANDARD DEGREES TO DECIMAL DEGREES!-->
									<xsl:variable name="latd" select="substring(substring-after(.,$transaction),7,2)"/>
									<xsl:variable name="latm" select="number(substring(substring-after(.,$transaction),9,2)) div 60"/>
									<xsl:variable name="lats" select="number(concat(substring(substring-after(.,$transaction),11,2),'.',substring(substring-after(.,$transaction),13,1))) div 3600"/>
									<xsl:variable name="latD" select = "concat((number($latd) + number($latm) + number($lats)) ,'.')" />
									<xsl:variable name="latFormat" select = "substring(concat(number(concat(substring('000000',1, 6 - string-length(substring-before($latD,'.'))),substring-before($latD,'.'),substring-after($latD,'.'))),'00000000'),1,7)" />									
									<xsl:variable name="eq" select="substring(substring-after(.,$transaction),14,1)"/>
									<xsl:variable name="longd" select="substring(substring-after(.,$transaction),15,3)"/>
									<xsl:variable name="longm" select="number(substring(substring-after(.,$transaction),18,2)) div 60"/>
									<xsl:variable name="longs" select="number(concat(substring(substring-after(.,$transaction),20,2),'.',substring(substring-after(.,$transaction),22,1))) div 3600 "/>
									<xsl:variable name="longD" select = "substring-after((number($longm) + number($longs)),'.')" />
									<xsl:variable name="longFormat" select = "substring(concat($eq, $longd,$longD, '000000000'),1,9)" />
									<MessageStatus>0</MessageStatus> 
									<rawMessage>
										<xsl:value-of select = "substring(.,1,9)"/>
										<xsl:value-of select="$latFormat"/>
										<xsl:value-of select="$longFormat"/>
										<xsl:value-of select = "substring(.,26)"/>
									</rawMessage>
								</xsl:when>
								<xsl:otherwise>
									<MessageStatus>0</MessageStatus>
									<rawMessage>
										<xsl:value-of select="/"/>
									</rawMessage>
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