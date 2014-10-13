<xsl:stylesheet xmlns:xsl = "http://www.w3.org/1999/XSL/Transform" version = "1.0" > 
				<xsl:output method = "xml" standalone ="yes" encoding="UTF-8"/> 
					<!--When transmitting coordinates to the CAD, the CAD expects Sexagesimal coordinates if the device transmits as Decimal use this transformation!-->
					<xsl:variable name = "transaction" select ="substring(.,2,2)"/>
					<xsl:variable name = "fromDecimal" select = "'true'"/>
					<xsl:template match="/">
					<GPSTransform>
						<xsl:choose>
						<xsl:when test="$transaction='LN'">
							<xsl:choose>
								<xsl:when test="$fromDecimal='true'">
									<xsl:variable name="lat" select="substring(substring-after(.,$transaction),10,2)" />
									<xsl:variable name="latM" select="substring-before(number(concat('.',substring(substring-after(.,$transaction),12,7)))*60,'.')" />
									<xsl:variable name="latMS" select = "concat(substring('00',1,2-string-length(string($latM))),$latM)"/>
									<xsl:variable name="latS" select="substring-after(number(concat('.',substring(substring-after(.,$transaction),12,7)))*60,'.')" />
									<xsl:variable name="latSS" select ="concat(substring-before(number(concat('.',$latS))*60,'.'),substring-after(number(concat('.',$latS))*60,'.'))"/>
									<xsl:variable name="eq" select="substring(substring-after(.,$transaction),19,1)"/>
									<xsl:variable name="long" select="substring(substring-after(.,$transaction),20,3)" />
									<xsl:variable name="longM" select="substring-before(number(concat('.',substring(substring-after(.,$transaction),23,7))*60),'.')" />
									<xsl:variable name="longMS" select = "concat(substring('00',1,2-string-length(string($longM))),$longM)"/>
									<xsl:variable name="longS" select="concat('.',substring-after(number(concat('.',substring(substring-after(.,$transaction),23,7))*60),'.'))" />
									<xsl:variable name="longSS" select = "concat(substring-before(number($longS)*60,'.'),substring-after(number($longS)*60,'.'))"/>
									<MessageStatus>0</MessageStatus>
									<rawMessage>
										<xsl:value-of select = "substring(.,1,12)"/>
										<xsl:value-of select="substring(concat($lat,$latMS,$latSS,'0000000000'),1,9)"/>
										<xsl:value-of select="substring(concat($eq,$long,$longMS,$longSS,'0000000000'),1,11)"/>
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
								<xsl:when test="$fromDecimal='true'">
									<xsl:variable name="lat" select="substring(substring-after(.,$transaction),7,2)" />
									<xsl:variable name="latM" select="substring-before(number(concat('.',substring(substring-after(.,$transaction),9,5)))*60,'.')" />
									<xsl:variable name="latMS" select = "concat(substring('00',1,2-string-length(string($latM))),$latM)"/>
									<xsl:variable name="latS" select="substring-after(number(concat('.',substring(substring-after(.,$transaction),9,5)))*60,'.')" />
									<xsl:variable name="latSS" select ="concat(substring-before(number(concat('.',$latS))*60,'.'),substring-after(number(concat('.',$latS))*60,'.'))"/>
									<xsl:variable name="eq" select="substring(substring-after(.,$transaction),14,1)"/>
									<xsl:variable name="long" select="substring(substring-after(.,$transaction),15,3)" />
									<xsl:variable name="longM" select="substring-before(number(concat('.',substring(substring-after(.,$transaction),18,5))*60),'.')" />
									<xsl:variable name="longMS" select = "concat(substring('00',1,2-string-length(string($longM))),$longM)"/>
									<xsl:variable name="longS" select="concat('.',substring-after(number(concat('.',substring(substring-after(.,$transaction),18,5))*60),'.'))" />
									<xsl:variable name="longSS" select = "concat(substring-before(number($longS)*60,'.'),substring-after(number($longS)*60,'.'))"/>
									<MessageStatus>0</MessageStatus>
									<rawMessage>
										<xsl:value-of select = "substring(.,1,9)"/>
										<xsl:value-of select="substring(concat($lat,$latMS,$latSS,'0000000'),1,7)"/>
										<xsl:value-of select="substring(concat($eq,$long,$longMS,$longSS,'000000000'),1,9)"/>
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