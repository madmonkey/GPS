<xsl:stylesheet xmlns:xsl = "http://www.w3.org/1999/XSL/Transform" version = "1.0" >
				<xsl:output method = "xml" standalone ="yes"/> 
					<xsl:template match="/">
					<GPSTransform>
							<!-- THIS HAS BEEN TESTED WITH A CADV MCS CONFIGURATION, CAD400 WOULD REQUIRE A BIT
							DIFFERENT LOGIC TO FIND THE UNIT ASSOCIATED WITH IT!-->
							<!-- What is our message name for this transaction ex: HGPS!-->
							<xsl:variable name="mapTrans"  select ="substring(substring-before(.,'/#*x02;'),string-length(substring-before(.,'/#*x02;'))-3,4)"/>
							<!-- What is kind of message ex: TAIP, NMEA!-->
							<xsl:variable name="transName"  select ="substring(substring-after(substring-after(.,$mapTrans),'/#*x02;'),1,4)"/>
							<!-- What is the sub-type of message ex: PV, LN, GGA !-->
							<xsl:variable name="transaction" select="substring(substring-after(substring-after(.,$transName),'R'),1,2)"/>
							<!-- What is the destination, according to message !-->
							<xsl:variable name="serverName"  select ="substring(substring-before(.,$mapTrans),string-length(substring-before(.,$mapTrans))-28,4)"/>
							<!-- Who is this message from!-->
							<xsl:variable name="unitID"  select ="substring(substring-before(.,$serverName),string-length(substring-before(.,$serverName))-7,4)"/>
							<xsl:choose>
								<xsl:when test = "$transaction ='PV'" >
									<MessageStatus>0</MessageStatus>
									<rawMessage>
									
									<xsl:value-of select="concat('/#*x02;','')"/>
									<xsl:value-of select ="substring(concat('CAD4Car','                    '),1,20)"/>
									<xsl:variable name="UnitID" select = "substring-before(.,$transaction)"/>
									<xsl:value-of select ="substring(concat($unitID,'                '),1,16)"/>
									<xsl:value-of select ="substring(concat('U702UNIT','            '),1,12)"/>
									<xsl:value-of select ="substring(concat('P',$unitID,'                                '),1,32)"/>
									<xsl:value-of select = "concat('ITEM','    ','ADD','     ','N')"/>
									<!-- I believe the +/- is counted as a positional space for the latitude !-->
									<xsl:variable name="lat" select="concat(concat(substring(substring-after(.,$transaction),7,2),'.'),substring(substring-after(.,$transaction),9,4))"/>
									<xsl:variable name="latD" select = "substring(concat(substring(number($lat) * 360000,1,10),'          '),1,10)" />
									<xsl:value-of select="$latD"/>
									<xsl:variable name="long" select="concat(concat(substring(substring-after(.,$transaction),15,4),'.'),substring(substring-after(.,$transaction),19,4))"/>
									<!--truncate return value to the first 8 characters, seems to like that better !-->
									<xsl:variable name="longD" select = "substring(concat(substring(number($long) * 360000,1,8),'          '),1,10)" />
									<xsl:value-of select="concat($longD, '                                                                       AVAIL')"/>
									<xsl:value-of select="concat('                                                                                                                                                                                                                                                           ','/#*x03;')"/>
									</rawMessage>
									<Type>0</Type>
								</xsl:when>
								<xsl:otherwise>
									<MessageStatus>3</MessageStatus>
									<rawMessage>Not currently processing this transaction type</rawMessage>
									<Type>0</Type>
								</xsl:otherwise>
							</xsl:choose>
					</GPSTransform>
					</xsl:template>
				</xsl:stylesheet>