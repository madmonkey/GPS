<xsl:stylesheet xmlns:xsl = "http://www.w3.org/1999/XSL/Transform" version = "1.0" >
				<xsl:output method = "xml" standalone ="yes"/> 
					<xsl:template match="/">
					<GPSTransform>
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
							
							<!--ASSUMPTION IS THAT WE ARE RECEIVING DECIMAL COORDINATES-->

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
										<!--CONVERT FROM DECIMAL DEGREES TO DECIMAL SECONDS!-->
										<xsl:variable name="latd" select="substring(substring-after(.,$transaction),7,2)"/>
										<xsl:variable name="latSeconds" select = "concat(number(concat($latd, '.', substring(substring-after(.,$transaction),9,5))) *3600,'.')"/>
										<!--ADDITIONAL ML PHENOMENION - DECIMAL MUST BE LAST 2 POSTIONS, SECONDS THE FIRST 6 !-->
										<xsl:variable name="latFormat" select ="concat(substring('000000',1, 6 - string-length(substring-before($latSeconds,'.'))),substring-before($latSeconds,'.'),substring(concat(substring-after($latSeconds,'.'),'00'),1,2))"/>
										<!--MAPLITE ONLY USES THE FIRST 8 CHARACTERS, ALTHOUGH THE FIELD IS SET-UP FOR 10 -->
										<xsl:value-of select = "concat($latFormat,substring('          ',1,10 - string-length($latFormat)))"/>
										<!--CONVERT FROM STANDARD DEGREES TO DECIMAL DEGREES TO DECIMAL SECONDS!-->
										<xsl:variable name="longd" select="substring(substring-after(.,$transaction),15,3)"/>
										<xsl:variable name="longSeconds" select="concat(number(concat($longd,'.',substring(substring-after(.,$transaction),18,5))) *3600,'.')"/>
										<!--ADDITIONAL ML PHENOMENION - DECIMAL MUST BE LAST 2 POSTIONS, SECONDS THE FIRST 6 !-->
										<xsl:variable name="longFormat" select ="concat(substring('000000',1, 6 - string-length(substring-before($longSeconds,'.'))),substring-before($longSeconds,'.'),substring(concat(substring-after($longSeconds,'.'),'00'),1,2))"/>
										<xsl:value-of select = "concat($longFormat,substring('          ',1,10 - string-length($longFormat)))"/>										<xsl:value-of select="'/#*x03;'"/>
										<!--ABRIDGED MAPLITE MESSAGE
										<xsl:value-of select ="'                                                                       AVAIL')"/>
										<xsl:value-of select="concat('                                                                                                                                                                                                                                                           ','/#*x03;')"/>
										!-->
									</rawMessage>
									<Type>0</Type>
								</xsl:when>
								<xsl:when test = "$transaction ='LN'" >
									<MessageStatus>0</MessageStatus>
									<rawMessage>
										<xsl:value-of select="concat('/#*x02;','')"/>
										<xsl:value-of select ="substring(concat('CAD4Car','                    '),1,20)"/>
										<xsl:variable name="UnitID" select = "substring-before(.,$transaction)"/>
										<xsl:value-of select ="substring(concat($unitID,'                '),1,16)"/>
										<xsl:value-of select ="substring(concat('U702UNIT','            '),1,12)"/>
										<xsl:value-of select ="substring(concat('P',$unitID,'                                '),1,32)"/>
										<xsl:value-of select = "concat('ITEM','    ','ADD','     ','N')"/>
										<!--CONVERT FROM DECIMAL DEGREES TO DECIMAL SECONDS!-->
										<xsl:variable name="latd" select="substring(substring-after(.,$transaction),10,2)"/>
										<xsl:variable name="latSeconds" select = "concat(number(concat($latd,'.',substring(substring-after(.,$transaction),12,7))) * 3600,'.')" />
										<xsl:variable name="latFormat" select ="concat(substring('000000',1, 6 - string-length(substring-before($latSeconds,'.'))),substring-before($latSeconds,'.'),substring(concat(substring-after($latSeconds,'.'),'00'),1,2))"/>
										<!--MAPLITE ONLY USES THE FIRST 8 CHARACTERS HENCE THE DOUBLE SUBSTRING!-->
										<xsl:value-of select = "concat($latFormat,substring('          ',1,10 - string-length($latFormat)))"/>										<!--CONVERT FROM STANDARD DEGREES TO DECIMAL DEGREES TO DECIMAL SECONDS!-->
										<xsl:variable name="longd" select="substring(substring-after(.,$transaction),20,3)"/>
										<xsl:variable name="longSeconds" select="concat(number(concat($longd,'.',substring(substring-after(.,$transaction),23,7))) *3600,'.')"/>
										<!--ADDITIONAL ML PHENOMENION - DECIMAL MUST BE LAST 2 POSTIONS, SECONDS THE FIRST 6 !-->
										<xsl:variable name="longFormat" select ="concat(substring('000000',1, 6 - string-length(substring-before($longSeconds,'.'))),substring-before($longSeconds,'.'),substring(concat(substring-after($longSeconds,'.'),'00'),1,2))"/>
										<!--MAPLITE ONLY USES THE FIRST 8 CHARACTERS, THOUGH THE FIELD IS SET-UP FOR 10 -->
										<xsl:value-of select = "concat($longFormat,substring('          ',1,10 - string-length($longFormat)))"/>										<xsl:value-of select="'/#*x03;'"/>
										<!--ABRIDGED MAPLITE MESSAGE
										<xsl:value-of select ="'                                                                       AVAIL')"/>
										<xsl:value-of select="concat('                                                                                                                                                                                                                                                           ','/#*x03;')"/>
										!-->
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