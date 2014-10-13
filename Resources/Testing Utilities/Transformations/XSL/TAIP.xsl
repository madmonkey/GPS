<xsl:stylesheet xmlns:xsl = "http://www.w3.org/1999/XSL/Transform" version = "1.0" > 
				<xsl:output method = "xml" standalone ="yes" encoding="UTF-8"/> 
					<!--When transmitting coordinates to the CAD, (CADV or CAD400) expects Decimal (DD) coordinates if the device transmits as Sexagesimal (DMS) use this transformation!-->
					<xsl:variable name = "transaction" select ="substring(.,2,2)"/>
					<xsl:variable name = "toDecimal" select = "'true'"/>
					<xsl:template match="/">
					<GPSTransform>
						<xsl:choose>
						<xsl:when test="$transaction='LN'">
							<xsl:choose>
								<xsl:when test="$toDecimal='true'">
									<xsl:variable name="lat" select="substring(substring-after(.,$transaction),10,2)" />
									<xsl:variable name="latM" select="number(substring(substring-after(.,$transaction),12,2)) div 60" />
									<xsl:variable name="latS" select="number(concat(substring(substring-after(.,$transaction),14,2),'.',substring(substring-after(.,$transaction),16,3))) div 3600 " />
									<xsl:variable name="latValue" select ="($latM + $latS)"/>
									<xsl:variable name="latFormat" select ="concat(concat(substring('00',1,2-string-length(string($lat))),$lat), substring-after($latValue,'.'))"/>
									<xsl:variable name="eq" select="substring(substring-after(.,$transaction),19,1)"/>
									<xsl:variable name="long" select="substring(substring-after(.,$transaction),20,3)" />
									<xsl:variable name="longM" select="number(substring(substring-after(.,$transaction),23,2)) div 60" />
									<xsl:variable name="longS" select = "number(concat(substring(substring-after(.,$transaction),25,2),'.',substring(substring-after(.,$transaction),27,3)))  div 3600"/>
									<xsl:variable name="longValue" select ="($longM + $longS)"/>
									<xsl:variable name="longFormat" select ="concat(concat(substring('000',1,3-string-length(string($long))),$long), substring-after($longValue,'.'))"/>
									<MessageStatus>0</MessageStatus>
									<rawMessage>
										<xsl:value-of select = "substring(.,1,12)"/>
										<xsl:value-of select="substring(concat($latFormat,'0000000000'),1,9)"/>
										<xsl:value-of select="substring(concat($eq,$longFormat,'0000000000'),1,11)"/>
										<xsl:value-of select = "substring(.,33)"/>
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
						<xsl:when test="$transaction='PV'">
							<xsl:choose>
								<xsl:when test="$toDecimal='true'">
									<xsl:variable name="lat" select="substring(substring-after(.,$transaction),7,2)" />
									<xsl:variable name="latM" select="number(substring(substring-after(.,$transaction),9,2)) div 60" />
									<xsl:variable name="latS" select="number(concat(substring(substring-after(.,$transaction),11,2),'.',substring(substring-after(.,$transaction),13,1))) div 3600 " />
									<xsl:variable name="latValue" select ="($latM + $latS)"/>
									<xsl:variable name="latFormat" select ="concat(concat(substring('00',1,2-string-length(string($lat))),$lat), substring-after($latValue,'.'))"/>
									<xsl:variable name="eq" select="substring(substring-after(.,$transaction),14,1)"/>
									<xsl:variable name="long" select="substring(substring-after(.,$transaction),15,3)" />
									<xsl:variable name="longM" select="number(substring(substring-after(.,$transaction),18,2)) div 60" />
									<xsl:variable name="longS" select = "number(concat(substring(substring-after(.,$transaction),20,2),'.',substring(substring-after(.,$transaction),22,1))) div 3600"/>
									<xsl:variable name="longValue" select ="($longM + $longS)"/>
									<xsl:variable name="longFormat" select ="concat(concat(substring('000',1,3-string-length(string($long))),$long), substring-after($longValue,'.'))"/>
									<MessageStatus>0</MessageStatus>
									<rawMessage>
										<xsl:value-of select = "substring(.,1,9)"/>
										<xsl:value-of select = "substring(concat($latFormat,'0000000'),1,7)"/>
										<xsl:value-of select="substring(concat($eq,$longFormat,'000000000'),1,9)"/>
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