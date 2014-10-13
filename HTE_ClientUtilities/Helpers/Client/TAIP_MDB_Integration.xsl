<xsl:stylesheet xmlns:xsl = "http://www.w3.org/1999/XSL/Transform" version = "1.0" > 
<xsl:output method = "xml" standalone ="yes" encoding="UTF-8"/> 
<xsl:template match="/">

<xsl:variable name = "transaction" select ="substring(.,2,2)"/>
<!--TAIP_MDB_Integration.xsl (c) SunGard HTE Inc-->
		
<!--TAIP STYLESHEET-->

<!--INCLUDE THE COORDINATE SYSTEM FROM YOUR INCOMING DATA DEVICE 
	CHOICES ARE: 
	'DMS' (Degrees Minutes Seconds), 
	'DDM' (Degrees Decimal Minutes) NMEA 0183 Standard,
	'DD'  (Decimal Degrees)
-->
<xsl:variable name = "fromSystem" select = "'DD'"/>

<!--WHAT COORDINATE SYSTEM DOES THE OUTPUT NEED TO BE 
	CHOICES ARE:
	'DD' (Decimal Degrees)
-->
<xsl:variable name =  "toSystem" select = "'DD'"/>

		<GPSTransform>
			<xsl:choose>
				<xsl:when test="$transaction='LN'">
					<xsl:choose>
						<!--PERFORMS CONVERSION SELECTING THE VALUES-->
						<xsl:when test="$fromSystem='DMS' and $toSystem='DD'">
							<xsl:variable name ="secs" select ="concat(substring(substring-after(.,$transaction),1,5),'.',substring(substring-after(.,$transaction),6,3))"/>
							<xsl:variable name="hm" select ="substring(substring-after(.,$transaction),9,1)"/>
							<xsl:variable name="lat" select="substring(substring-after(.,$transaction),10,2)" />
							<xsl:variable name="latM" select="number(substring(substring-after(.,$transaction),12,2)) div 60" />
							<xsl:variable name="latS" select="number(concat(substring(substring-after(.,$transaction),14,2),'.',substring(substring-after(.,$transaction),16,3))) div 3600 " />
							<xsl:variable name="latValue" select ="($latM + $latS)"/>
							<xsl:variable name="latFormat" select ="concat(concat(substring('00',1,2-string-length(string($lat))),$lat), '.', substring-after($latValue,'.'))"/>
							<xsl:variable name="eq" select="substring(substring-after(.,$transaction),19,1)"/>
							<xsl:variable name="long" select="substring(substring-after(.,$transaction),20,3)" />
							<xsl:variable name="longM" select="number(substring(substring-after(.,$transaction),23,2)) div 60" />
							<xsl:variable name="longS" select = "number(concat(substring(substring-after(.,$transaction),25,2),'.',substring(substring-after(.,$transaction),27,3)))  div 3600"/>
							<xsl:variable name="longValue" select ="($longM + $longS)"/>
							<xsl:variable name="longFormat" select ="concat(concat(substring('000',1,3-string-length(string($long))),$long), '.', substring-after($longValue,'.'))"/>
							<xsl:variable name="hhmmss">
								<xsl:variable name="hh" select="floor($secs div 3600)"/>
								<xsl:variable name="mm" select="floor($secs div 60) mod 60"/>
								<xsl:variable name="ss" select="$secs mod 60"/>
								<xsl:if test="$hh &lt; 10">
									<xsl:text>0</xsl:text>
								</xsl:if>
								<xsl:value-of select="$hh"/>
								<xsl:text>:</xsl:text>
								<xsl:if test="$mm &lt; 10">
									<xsl:text>0</xsl:text>
								</xsl:if>
								<xsl:value-of select="$mm"/>
								<xsl:text>:</xsl:text>
								<xsl:if test="$ss &lt; 10">
									<xsl:text>0</xsl:text>
								</xsl:if>
								<xsl:value-of select="$ss"/>
							</xsl:variable>
							<xsl:choose>
								<xsl:when test="number($latValue)!=0">
									<MessageStatus>0</MessageStatus>
									<rawMessage method="xml">
										<MDBGPS>
											<LAT><xsl:value-of select="concat($hm, $latFormat)"/></LAT>
											<LONG><xsl:value-of select ="concat($eq, $longFormat)"/></LONG>
											<TIMESTAMP>
												<xsl:value-of select="$hhmmss"/>
											</TIMESTAMP>
										</MDBGPS>
										<!--
										<xsl:value-of select = "substring(.,1,12)"/>
										<xsl:value-of select="substring(concat($latFormat,'0000000000'),1,9)"/>
										<xsl:value-of select="substring(concat($eq,$longFormat,'0000000000'),1,11)"/>
										<xsl:value-of select = "substring(.,33)"/>
										-->
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
						<!--PERFORMS CONVERSION SELECTING THE VALUES-->
						<xsl:when test="$fromSystem='DDM' and $toSystem='DD'">
							<xsl:variable name ="secs" select ="concat(substring(substring-after(.,$transaction),1,5),'.',substring(substring-after(.,$transaction),6,3))"/>
							<xsl:variable name="hm" select ="substring(substring-after(.,$transaction),9,1)"/>
							<xsl:variable name="lat" select="substring(substring-after(.,$transaction),10,2)" />
							<xsl:variable name="latM" select="number(substring(substring-after(.,$transaction),12,7)) div 60" />
							<xsl:variable name="latFormat" select ="concat(concat(substring('00',1,2-string-length(string($lat))),$lat), '.', substring-before($latM,'.'),substring-after($latM,'.'))"/>
							<xsl:variable name="eq" select="substring(substring-after(.,$transaction),19,1)"/>
							<xsl:variable name="long" select="substring(substring-after(.,$transaction),20,3)" />
							<xsl:variable name="longM" select="number(substring(substring-after(.,$transaction),23,7)) div 60" />
							<xsl:variable name="longFormat" select ="concat(concat(substring('000',1,3-string-length(string($long))),$long), '.', substring-before($longM,'.'),substring-after($longM,'.'))"/>
							<xsl:variable name="hhmmss">
								<xsl:variable name="hh" select="floor($secs div 3600)"/>
								<xsl:variable name="mm" select="floor($secs div 60) mod 60"/>
								<xsl:variable name="ss" select="$secs mod 60"/>
								<xsl:if test="$hh &lt; 10">
									<xsl:text>0</xsl:text>
								</xsl:if>
								<xsl:value-of select="$hh"/>
								<xsl:text>:</xsl:text>
								<xsl:if test="$mm &lt; 10">
									<xsl:text>0</xsl:text>
								</xsl:if>
								<xsl:value-of select="$mm"/>
								<xsl:text>:</xsl:text>
								<xsl:if test="$ss &lt; 10">
									<xsl:text>0</xsl:text>
								</xsl:if>
								<xsl:value-of select="$ss"/>
							</xsl:variable>
							<xsl:choose>
								<xsl:when test="number($latFormat)!=0">
									<MessageStatus>0</MessageStatus>
									<rawMessage method="xml">
										<MDBGPS>
											<LAT><xsl:value-of select="concat($hm, $latFormat)"/></LAT>
											<LONG><xsl:value-of select ="concat($eq, $longFormat)"/></LONG>
											<TIMESTAMP><xsl:value-of select="$hhmmss"/></TIMESTAMP>
										</MDBGPS>
										<!--
										<xsl:value-of select = "substring(.,1,12)"/>
										<xsl:value-of select="substring(concat($latFormat,'0000000000'),1,9)"/>
										<xsl:value-of select="substring(concat($eq,$longFormat,'0000000000'),1,11)"/>
										<xsl:value-of select = "substring(.,33)"/>
										-->
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
							<xsl:variable name ="secs" select ="concat(substring(substring-after(.,$transaction),1,5),'.',substring(substring-after(.,$transaction),6,3))"/>
							<xsl:variable name="hm" select ="substring(substring-after(.,$transaction),9,1)"/>
							<xsl:variable name="lat" select="substring(substring-after(.,$transaction),10,2)" />
							<xsl:variable name="latM" select="substring(substring-after(.,$transaction),12,7)" />
							<xsl:variable name="latFormat" select ="concat(concat(substring('00',1,2-string-length(string($lat))),$lat), '.', $latM)"/>
							<xsl:variable name="eq" select="substring(substring-after(.,$transaction),19,1)"/>
							<xsl:variable name="long" select="substring(substring-after(.,$transaction),20,3)" />
							<xsl:variable name="longM" select="substring(substring-after(.,$transaction),23,7)" />
							<xsl:variable name="longFormat" select ="concat(concat(substring('000',1,3-string-length(string($long))),$long), '.', $longM)"/>
							<xsl:variable name="hhmmss">
								<xsl:variable name="hh" select="floor($secs div 3600)"/>
								<xsl:variable name="mm" select="floor($secs div 60) mod 60"/>
								<xsl:variable name="ss" select="$secs mod 60"/>
								<xsl:if test="$hh &lt; 10">
									<xsl:text>0</xsl:text>
								</xsl:if>
								<xsl:value-of select="$hh"/>
								<xsl:text>:</xsl:text>
								<xsl:if test="$mm &lt; 10">
									<xsl:text>0</xsl:text>
								</xsl:if>
								<xsl:value-of select="$mm"/>
								<xsl:text>:</xsl:text>
								<xsl:if test="$ss &lt; 10">
									<xsl:text>0</xsl:text>
								</xsl:if>
								<xsl:value-of select="$ss"/>
							</xsl:variable>
							<!--Assume Decimal Degrees-->
							<xsl:choose>
								<xsl:when test ="number(substring(substring-after(.,$transaction),10,9))!=0">
									<MessageStatus>0</MessageStatus>
									<rawMessage method="xml">
										<MDBGPS>
											<LAT><xsl:value-of select="concat($hm, $latFormat)"/></LAT>
											<LONG><xsl:value-of select ="concat($eq, $longFormat)"/></LONG>
											<TIMESTAMP><xsl:value-of select="$hhmmss"/></TIMESTAMP>
												
										</MDBGPS>
									</rawMessage>
									<Type>0</Type>
								</xsl:when>
								<xsl:otherwise>
									<MessageStatus>3</MessageStatus>
									<rawMessage>NO VALID DATA</rawMessage>
									<Type>0</Type>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<xsl:when test="$transaction='PV'">
					<xsl:choose>
						<!--PERFORMS CONVERSION SELECTING THE VALUES-->
						<xsl:when test="$fromSystem='DMS' and $toSystem ='DD'">
							<xsl:variable name ="secs" select ="substring(substring-after(.,$transaction),1,5)"/>
							<xsl:variable name="hm" select ="substring(substring-after(.,$transaction),6,1)"/>
							<xsl:variable name="lat" select="substring(substring-after(.,$transaction),7,2)" />
							<xsl:variable name="latM" select="number(substring(substring-after(.,$transaction),9,2)) div 60" />
							<xsl:variable name="latS" select="number(concat(substring(substring-after(.,$transaction),11,2),'.',substring(substring-after(.,$transaction),13,1))) div 3600 " />
							<xsl:variable name="latValue" select ="($latM + $latS)"/>
							<xsl:variable name="latFormat" select ="concat(concat(substring('00',1,2-string-length(string($lat))),$lat), '.', substring-after($latValue,'.'))"/>
							<xsl:variable name="eq" select="substring(substring-after(.,$transaction),14,1)"/>
							<xsl:variable name="long" select="substring(substring-after(.,$transaction),15,3)" />
							<xsl:variable name="longM" select="number(substring(substring-after(.,$transaction),18,2)) div 60" />
							<xsl:variable name="longS" select = "number(concat(substring(substring-after(.,$transaction),20,2),'.',substring(substring-after(.,$transaction),22,1))) div 3600"/>
							<xsl:variable name="longValue" select ="($longM + $longS)"/>
							<xsl:variable name="longFormat" select ="concat(concat(substring('000',1,3-string-length(string($long))),$long), '.', substring-after($longValue,'.'))"/>
							<xsl:variable name="hhmmss">
								<xsl:variable name="hh" select="floor($secs div 3600)"/>
								<xsl:variable name="mm" select="floor($secs div 60) mod 60"/>
								<xsl:variable name="ss" select="$secs mod 60"/>
								<xsl:if test="$hh &lt; 10">
									<xsl:text>0</xsl:text>
								</xsl:if>
								<xsl:value-of select="$hh"/>
								<xsl:text>:</xsl:text>
								<xsl:if test="$mm &lt; 10">
									<xsl:text>0</xsl:text>
								</xsl:if>
								<xsl:value-of select="$mm"/>
								<xsl:text>:</xsl:text>
								<xsl:if test="$ss &lt; 10">
									<xsl:text>0</xsl:text>
								</xsl:if>
								<xsl:value-of select="$ss"/>
							</xsl:variable>
							<xsl:choose>
								<xsl:when test="number($latValue)!=0">
									<MessageStatus>0</MessageStatus>
									<rawMessage method="xml">
										<MDBGPS>
											<LAT><xsl:value-of select="concat($hm, $latFormat)"/></LAT>
											<LONG><xsl:value-of select ="concat($eq, $longFormat)"/></LONG>
											<TIMESTAMP><xsl:value-of select="$hhmmss"/></TIMESTAMP>
										</MDBGPS>
										<!--
										<xsl:value-of select = "substring(.,1,9)"/>
										<xsl:value-of select = "substring(concat($latFormat,'0000000'),1,7)"/>
										<xsl:value-of select="substring(concat($eq,$longFormat,'000000000'),1,9)"/>
										<xsl:value-of select = "substring(.,26)"/>
										-->
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
						<!--PERFORMS CONVERSION SELECTING THE VALUES-->
						<xsl:when test="$fromSystem='DDM' and $toSystem='DD'">
							<xsl:variable name ="secs" select ="substring(substring-after(.,$transaction),1,5)"/>
							<xsl:variable name="hm" select ="substring(substring-after(.,$transaction),6,1)"/>
							<xsl:variable name="lat" select="substring(substring-after(.,$transaction),7,2)" />
							<xsl:variable name="latM" select="number(substring(substring-after(.,$transaction),9,5)) div 60" />
							<xsl:variable name="latFormat" select ="concat(concat(substring('00',1,2-string-length(string($lat))),$lat), substring-before($latM,'.'),substring-after($latM,'.'))"/>
							<xsl:variable name="eq" select="substring(substring-after(.,$transaction),14,1)"/>
							<xsl:variable name="long" select="substring(substring-after(.,$transaction),15,3)" />
							<xsl:variable name="longM" select="number(substring(substring-after(.,$transaction),18,5)) div 60" />
							<xsl:variable name="longFormat" select ="concat(concat(substring('000',1,3-string-length(string($long))),$long), substring-before($longM,'.'),substring-after($longM,'.'))"/>
							<xsl:variable name="hhmmss">
								<xsl:variable name="hh" select="floor($secs div 3600)"/>
								<xsl:variable name="mm" select="floor($secs div 60) mod 60"/>
								<xsl:variable name="ss" select="$secs mod 60"/>
								<xsl:if test="$hh &lt; 10">
									<xsl:text>0</xsl:text>
								</xsl:if>
								<xsl:value-of select="$hh"/>
								<xsl:text>:</xsl:text>
								<xsl:if test="$mm &lt; 10">
									<xsl:text>0</xsl:text>
								</xsl:if>
								<xsl:value-of select="$mm"/>
								<xsl:text>:</xsl:text>
								<xsl:if test="$ss &lt; 10">
									<xsl:text>0</xsl:text>
								</xsl:if>
								<xsl:value-of select="$ss"/>
							</xsl:variable>
							<xsl:choose>
								<xsl:when test="number($latFormat)!=0">
									<MessageStatus>0</MessageStatus>
									<rawMessage method="xml">
										<MDBGPS>
											<LAT><xsl:value-of select="concat($hm, $latFormat)"/></LAT>
											<LONG><xsl:value-of select ="concat($eq, $longFormat)"/></LONG>
											<TIMESTAMP><xsl:value-of select="$hhmmss"/></TIMESTAMP>
										</MDBGPS>
										<!--
										<xsl:value-of select = "substring(.,1,9)"/>
										<xsl:value-of select = "substring(concat($latFormat,'0000000'),1,7)"/>
										<xsl:value-of select="substring(concat($eq,$longFormat,'000000000'),1,9)"/>
										<xsl:value-of select = "substring(.,26)"/>
										-->
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
							<!--Assume Decimal Degrees-->
							<xsl:variable name ="secs" select ="substring(substring-after(.,$transaction),1,5)"/>
							<xsl:variable name="hm" select ="substring(substring-after(.,$transaction),6,1)"/>
							<xsl:variable name="lat" select="substring(substring-after(.,$transaction),7,2)" />
							<xsl:variable name="latM" select="substring(substring-after(.,$transaction),9,5)" />
							<xsl:variable name="latFormat" select ="concat(concat(substring('00',1,2-string-length(string($lat))),$lat), '.', $latM)"/>
							<xsl:variable name="eq" select="substring(substring-after(.,$transaction),14,1)"/>
							<xsl:variable name="long" select="substring(substring-after(.,$transaction),15,3)" />
							<xsl:variable name="longM" select="substring(substring-after(.,$transaction),18,5)" />
							<xsl:variable name="longFormat" select ="concat(concat(substring('000',1,3-string-length(string($long))),$long), '.', $longM)"/>
							<xsl:variable name="hhmmss">
								<xsl:variable name="hh" select="floor($secs div 3600)"/>
								<xsl:variable name="mm" select="floor($secs div 60) mod 60"/>
								<xsl:variable name="ss" select="$secs mod 60"/>
								<xsl:if test="$hh &lt; 10">
									<xsl:text>0</xsl:text>
								</xsl:if>
								<xsl:value-of select="$hh"/>
								<xsl:text>:</xsl:text>
								<xsl:if test="$mm &lt; 10">
									<xsl:text>0</xsl:text>
								</xsl:if>
								<xsl:value-of select="$mm"/>
								<xsl:text>:</xsl:text>
								<xsl:if test="$ss &lt; 10">
									<xsl:text>0</xsl:text>
								</xsl:if>
								<xsl:value-of select="$ss"/>
							</xsl:variable>
							<xsl:choose>
								<xsl:when test ="number(substring(substring-after(.,$transaction),7,6))!=0">
									<MessageStatus>0</MessageStatus>
									<rawMessage method="xml">
										<MDBGPS>
											<LAT><xsl:value-of select="concat($hm, $latFormat)"/></LAT>
											<LONG><xsl:value-of select ="concat($eq, $longFormat)"/></LONG>
											<TIMESTAMP><xsl:value-of select="$hhmmss"/></TIMESTAMP>
										</MDBGPS>
									</rawMessage>
									<Type>0</Type>
								</xsl:when>
								<xsl:otherwise>
									<MessageStatus>3</MessageStatus>
									<rawMessage>NO VALID DATA</rawMessage>
									<Type>0</Type>
								</xsl:otherwise>
							</xsl:choose>
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