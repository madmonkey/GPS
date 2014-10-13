<xsl:stylesheet xmlns:xsl = "http://www.w3.org/1999/XSL/Transform" version = "1.0" > 
				<xsl:output method = "xml" standalone ="yes"/> 
					<xsl:template match="/">
					<GPSTransform>
							<xsl:variable name="transaction" select="substring(substring-after(.,'R'),1,2)"/>
							<xsl:variable name="unitID" select="'/@#{ID};'"/>
							<!--IF YOU LEAVE STAT BLANK IT WILL DEFAULT TO THE CURRENT COLOR FROM CAD IF NOT
							THE MATRIX IS AS FOLLOWS 00=BLACK, 01=NAVY BLUE,02=GREEN,03=TEAL,04=MAROON,05=PURPLE,06=BROWN,07=SILVER,08=GREY,09=BLUE,10=NEON GREEN,11=LT.BLUE,12=RED,13=HOT PINK,14=YELLOW,15=WHITE
							!-->
							<xsl:variable name="stat" select="'  '"/>
							<xsl:choose>
								<xsl:when test = "$transaction ='PV'" >
									<MessageStatus>0</MessageStatus>
									<rawMessage>
										<xsl:value-of select="concat('/#*x02;','')"/>
										<xsl:value-of select ="substring(concat('CAD4Car','                    '),1,20)"/>
										<xsl:value-of select ="substring(concat($unitID,'                '),1,16)"/>
										<xsl:value-of select ="substring(concat('U7',$stat,'UNIT','            '),1,12)"/>
										<xsl:value-of select ="substring(concat('P',unitID,'                                '),1,32)"/>
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
										<xsl:value-of select = "concat($longFormat,substring('          ',1,10 - string-length($longFormat)))"/>
										<xsl:value-of select="'/#*x03;'"/>
										<!--ABRIDGED VERSION FOR SMALLER MAPLITE MESSAGE 
										<xsl:value-of select="('                                                                       AVAIL')"/>
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
										<xsl:value-of select ="substring(concat($unitID,'                '),1,16)"/>
										<xsl:value-of select ="substring(concat('U7',$stat,'UNIT','            '),1,12)"/>
										<xsl:value-of select ="substring(concat('P',$unitID,'                                '),1,32)"/>
										<xsl:value-of select = "concat('ITEM','    ','ADD','     ','N')"/>
										<!--CONVERT FROM DECIMAL DEGREES TO DECIMAL SECONDS!-->
										<xsl:variable name="latd" select="substring(substring-after(.,$transaction),10,2)"/>
										<xsl:variable name="latSeconds" select = "concat(number(concat($latd, '.', substring(substring-after(.,$transaction),12,7))) *3600,'.')"/>
										<!--ADDITIONAL ML PHENOMENION - DECIMAL MUST BE LAST 2 POSTIONS, SECONDS THE FIRST 6 !-->
										<xsl:variable name="latFormat" select ="concat(substring('000000',1, 6 - string-length(substring-before($latSeconds,'.'))),substring-before($latSeconds,'.'),substring(concat(substring-after($latSeconds,'.'),'00'),1,2))"/>
										<!--MAPLITE ONLY USES THE FIRST 8 CHARACTERS, THOUGH THE FIELD IS SET-UP FOR 10 -->
										<xsl:value-of select = "concat($latFormat,substring('          ',1,10 - string-length($latFormat)))"/>
										<xsl:variable name="longd" select="substring(substring-after(.,$transaction),20,3)"/>
										<xsl:variable name="longSeconds" select="concat(number(concat($longd,'.',substring(substring-after(.,$transaction),23,7))) *3600,'.')"/>
										<!--ADDITIONAL ML PHENOMENION - DECIMAL MUST BE LAST 2 POSTIONS, SECONDS THE FIRST 6 !-->
										<xsl:variable name="longFormat" select ="concat(substring('000000',1, 6 - string-length(substring-before($longSeconds,'.'))),substring-before($longSeconds,'.'),substring(concat(substring-after($longSeconds,'.'),'00'),1,2))"/>
										<!--MAPLITE ONLY USES THE FIRST 8 CHARACTERS, THOUGH THE FIELD IS SET-UP FOR 10 -->
										<xsl:value-of select = "concat($longFormat,substring('          ',1,10 - string-length($longFormat)))"/>
										<xsl:value-of select="'/#*x03;'"/>
										<!--ABRIDGED VERSION FOR SMALLER MAPLITE MESSAGE 
										<xsl:value-of select="('                                                                       AVAIL')"/>
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