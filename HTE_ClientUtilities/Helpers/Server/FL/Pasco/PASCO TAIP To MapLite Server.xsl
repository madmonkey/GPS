<xsl:stylesheet xmlns:xsl = "http://www.w3.org/1999/XSL/Transform" version = "1.0" >
<!--TAIP To MAPLITE SERVER.xsl (c) SunGard HTE Inc-->
		
<!--TAIP TO MAPLITE SERVER (DECIMAL DEGREES) STYLESHEET STANDARD (SINCE CAD(s) ACCEPT ONLY DECIMAL DEGREES)--> 
				<xsl:output method = "xml" standalone ="yes"/> 
					<xsl:template match="/">
					<GPSTransform>
							<xsl:variable name="transaction" select="substring(substring-after(.,'R'),1,2)"/>
							<!--modified to parse unit from DATA-STREAM rather than look-up based on unit-configuration -->
							<!--tested with/without checksum, with/without ID, and intermixed -->
							<!--EXPECT COORDINATES IN DECIMAL DEGREES TO PERFORM THE CONVERSION -->
							<xsl:variable name="unitID"><![CDATA[/@#{ENTITYDEV};]]></xsl:variable>
            
            <xsl:variable name="aliasID"><![CDATA[/@#{ENTITYUNIT};]]></xsl:variable>
            
							<xsl:variable name="status"><![CDATA[/@#{ENTITYSTAT};]]></xsl:variable>
							<!--IF YOU LEAVE STAT BLANK IT WILL DEFAULT TO THE CURRENT COLOR FROM CAD IF NOT
							THE MATRIX IS AS FOLLOWS 00=BLACK, 01=NAVY BLUE,02=GREEN,03=TEAL,04=MAROON,05=PURPLE,06=BROWN,07=SILVER,08=GREY,09=BLUE,10=NEON GREEN,11=LT.BLUE,12=RED,13=HOT PINK,14=YELLOW,15=WHITE
							!-->

            <xsl:variable name="AddTitle" select="'Y'"/>
            <xsl:variable name="stat" select="'  '"/>
							<xsl:choose>
								<xsl:when test = "$transaction ='PV'" >
									<MessageStatus>0</MessageStatus>
									<rawMessage>
										<xsl:value-of select ="substring(concat('CAD4Car','                    '),1,20)"/>
										<xsl:value-of select ="substring(concat($unitID,'                '),1,16)"/>
										<xsl:value-of select ="substring(concat('U7',$stat,'UNIT','            '),1,12)"/>

                    <xsl:choose>
                        <xsl:when test = "$AddTitle ='Y'" >                      
                        <xsl:variable name ="aliasIDM" select ="substring-before(substring-after(.,'TITLE='),';')"/>
                        <xsl:value-of select ="substring(concat('P',$aliasIDM,'                                '),1,32)"/>                      
                    </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select ="substring(concat('P',$aliasID,'                                '),1,32)"/>
                      </xsl:otherwise>

                    </xsl:choose>
                    
										
										<xsl:value-of select = "concat('ITEM','    ','ADD','     ')"/>
										<xsl:choose><xsl:when test = "substring(substring-after(.,$transaction),6,1)='+'">N</xsl:when><xsl:otherwise>S</xsl:otherwise></xsl:choose> 
										<!--LATITUDE: 
											WHAT: Convert from  decimal degrees (###.##### with decimal removed) 
												to decimal seconds (######.## with decimal removed). 

											HOW: Since we want two decimal places with decimal removed we calculate hundredths of seconds instead of seconds. 
												This moves the tenths and hundreths place in front of the decimal place. We only want the whole part of the result,
												and we want exactly eight digits, so we do the following: prepend 8 zeros, append a decimal point,  take everything
												up to the decimal point and from that take the right 8 digits. And, since maplite uses 8 digits, but expects 
												a 10 character field we append 2 spaces.
										!-->
										<xsl:variable name="latDegWhole" select="substring(substring-after(.,$transaction),7,2)"/>
										<xsl:variable name="latDegPart" select="substring(substring-after(.,$transaction),9,5)"/>
										<xsl:variable name="latHundredthsOfSeconds" select="substring-before(concat('00000000',concat($latDegWhole,'.',$latDegPart)*360000,'.'),'.')"/>
										<xsl:value-of  select="concat(substring($latHundredthsOfSeconds, string-length($latHundredthsOfSeconds) - 7, 8),'  ')"/>
										<!--LONGITUDE: 
											WHAT: Convert from  decimal degrees (###.##### with decimal removed) 
												to decimal seconds (######.## with decimal removed). 

											HOW: Since we want two decimal places with decimal removed we calculate hundredths of seconds instead of seconds. 
												This moves the tenths and hundreths place in front of the decimal place. We only want the whole part of the result,
												and we want exactly eight digits, so we do the following: prepend 8 zeros, append a decimal point,  take everything
												up to the decimal point and from that take the right 8 digits. And, since maplite uses 8 digits, but expects 
												a 10 character field we append 2 spaces.
										!-->
										<xsl:variable name="longDegWhole" select="substring(substring-after(.,$transaction),15,3)"/>
										<xsl:variable name="longDegPart" select="substring(substring-after(.,$transaction),18,5)"/>
										<xsl:variable name="longHundredthsOfSeconds" select="substring-before(concat('00000000',concat($longDegWhole,'.',$longDegPart)*360000,'.'),'.')"/>
										<xsl:value-of  select="concat(substring($longHundredthsOfSeconds, string-length($longHundredthsOfSeconds) - 7, 8),'  ')"/>
										
										<!--COMMENT LINE BELOW TO IGNORE SMART-CLIENT STATUS CODE DESCRIPTION !-->
										<xsl:value-of select="concat('                                                                       ',$status)"/>
										
										<!--ABRIDGED VERSION FOR SMALLER MAPLITE MESSAGE 
										<xsl:value-of select="concat('                                                                                                                                                                                                                                                           ','/#*x03;')"/>
										!-->
									</rawMessage>
									<Type>2</Type>
								</xsl:when>
								<xsl:when test = "$transaction ='LN'" >
									<MessageStatus>0</MessageStatus>
									<rawMessage>
										<!--Added MapLite as a default configuration use message type for tag(s)
										<xsl:value-of select="concat('/#*x02;','')"/>
										!-->
										<xsl:value-of select ="substring(concat('CAD4Car','                    '),1,20)"/>
										<xsl:value-of select ="substring(concat($unitID,'                '),1,16)"/>
										<xsl:value-of select ="substring(concat('U7',$stat,'UNIT','            '),1,12)"/>
										<xsl:value-of select ="substring(concat('P',$aliasID,'                                '),1,32)"/>
										<xsl:value-of select = "concat('ITEM','    ','ADD','     ')"/>
										<xsl:choose><xsl:when test = "substring(substring-after(.,$transaction),9,1)='+'">N</xsl:when><xsl:otherwise>S</xsl:otherwise></xsl:choose> 
										<!--LATITUDE: 
											WHAT: Convert from  decimal degrees (###.##### with decimal removed) 
												to decimal seconds (######.## with decimal removed). 

											HOW: Since we want two decimal places with decimal removed we calculate hundredths of seconds instead of seconds. 
												This moves the tenths and hundreths place in front of the decimal place. We only want the whole part of the result,
												and we want exactly eight digits, so we do the following: prepend 8 zeros, append a decimal point,  take everything
												up to the decimal point and from that take the right 8 digits. And, since maplite uses 8 digits, but expects 
												a 10 character field we append 2 spaces.
										!-->
										<xsl:variable name="latDegWhole" select="substring(substring-after(.,$transaction),10,2)"/>
										<xsl:variable name="latDegPart" select="substring(substring-after(.,$transaction),12,7)"/>
										<xsl:variable name="latHundredthsOfSeconds" select="substring-before(concat('00000000',concat($latDegWhole,'.',$latDegPart)*360000,'.'),'.')"/>
										<xsl:value-of  select="concat(substring($latHundredthsOfSeconds, string-length($latHundredthsOfSeconds) - 7, 8),'  ')"/>
										<!--LONGITUDE: 
											WHAT: Convert from  decimal degrees (###.##### with decimal removed) 
												to decimal seconds (######.## with decimal removed). 

											HOW: Since we want two decimal places with decimal removed we calculate hundredths of seconds instead of seconds. 
												This moves the tenths and hundreths place in front of the decimal place. We only want the whole part of the result,
												and we want exactly eight digits, so we do the following: prepend 8 zeros, append a decimal point,  take everything
												up to the decimal point and from that take the right 8 digits. And, since maplite uses 8 digits, but expects 
												a 10 character field we append 2 spaces.
										!-->
										<xsl:variable name="longDegWhole" select="substring(substring-after(.,$transaction),20,3)"/>
										<xsl:variable name="longDegPart" select="substring(substring-after(.,$transaction),23,7)"/>
										<xsl:variable name="longHundredthsOfSeconds" select="substring-before(concat('00000000',concat($longDegWhole,'.',$longDegPart)*360000,'.'),'.')"/>
										<xsl:value-of  select="concat(substring($longHundredthsOfSeconds, string-length($longHundredthsOfSeconds) - 7, 8),'  ')"/>
										<!--Added MapLite as a default configuration use message type for tag(s)
										<xsl:value-of select="'/#*x03;'"/>
										!-->

										<!--COMMENT LINE BELOW TO IGNORE SMART-CLIENT STATUS CODE DESCRIPTION !-->
										<xsl:value-of select="concat('                                                                       ',$status)"/>
										
										<!--ABRIDGED VERSION FOR SMALLER MAPLITE MESSAGE 
										<xsl:value-of select="concat('                                                                                                                                                                                                                                                           ','/#*x03;')"/>
										!-->
									</rawMessage>
									<Type>2</Type>
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