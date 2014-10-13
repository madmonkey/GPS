<xsl:stylesheet xmlns:xsl = "http://www.w3.org/1999/XSL/Transform" version = "1.0" > 
<xsl:output method = "xml" standalone ="yes" encoding="UTF-8"/> 
	<!--NMEA VALIDATE To TAIP PV SERVER.xsl (c) SunGard HTE Inc-->
		
	<!--NMEA VALIDATE TO TAIP PV SERVER STYLESHEET-->

		<!--INCLUDE THE COORDINATE SYSTEM FROM YOUR INCOMING DATA DEVICE 
			CHOICES ARE: 
			'DMS' (Degrees Minutes Seconds), 
			'DDM' (Degrees Decimal Minutes) NMEA 0183 Standard,
			'DD'  (Decimal Degrees)
		-->
		<xsl:variable name = "fromSystem" select = "'DDM'"/>

		<!--WHAT COORDINATE SYSTEM DOES THE OUTPUT NEED TO BE 
			CHOICES ARE:
			'DD' (Decimal Degrees)
		-->
		<xsl:variable name =  "toSystem" select = "'DD'"/>
		<!--THE NEW REQUIREMENT IS THAT WE WILL BE SENDING RAW TAIP COMPLIANT MESSAGES TO THE SERVER - TAG WITH ID-->
		<xsl:variable name="unitID"><![CDATA[/@#{ENTITYDEV};]]></xsl:variable>
		<xsl:variable name="aliasID"><![CDATA[/@#{ENTITYUNIT};]]></xsl:variable>
	<xsl:template match="/">
	<xsl:variable name="transaction" select="substring-before(.,',')"/>
	
	<GPSTransform>
	<xsl:choose>
		<!--GPGLL TRANSACTION-->
		<xsl:when test = "$transaction ='GPGLL'" >
			<xsl:variable name="lat" select="substring-before(substring-after(.,concat($transaction,',')),',')" />
			<xsl:variable name="latD" select="substring(substring-before($lat,'.'),1,2)" />
			<xsl:variable name="latM" select="substring(substring-before($lat,'.'),3)" />
			<xsl:variable name="latm" select="substring-after($lat,'.')" />
			<xsl:variable name="latMm" select = "concat(concat($latM,'.'),$latm)"/>
			<xsl:variable name="eq" select="substring(substring-after(.,$lat),2,1)" />
			<xsl:variable name="long" select="substring-before(substring(substring-after(.,$eq),2),',')" />
			<xsl:variable name="longD" select="substring(substring-before($long,'.'),1,3)" />
			<xsl:variable name="longM" select="substring(substring-before($long,'.'),4)" />
			<xsl:variable name="longm" select="substring-after($long,'.')" />
			<xsl:variable name="longMm" select = "concat(concat($longM,'.'),$longm)"/>
			<xsl:variable name="hm" select="substring-before(substring(substring-after(.,$long),2),',')" />
			<xsl:variable name="time" select="concat(substring-before(substring(substring-after(.,$hm),2),','),'00000000')" />
			<xsl:variable name="hours" select="substring($time,1,2)" />
			<xsl:variable name="mins" select="substring($time,3,2)" />
			<xsl:variable name="secs" select="substring($time,5,2)" />
			<xsl:variable name="active" select="substring-before(substring-after(substring(substring-after(.,$hm),2),','),',')" />
			<xsl:variable name="chksum" select="substring(substring-after(.,$active),2)" />
			<!--0 - processed, 1 - processed warn, 2 - processed error, 3 - error -->
			<xsl:choose>
				<xsl:when test ="string-length($latD)> 0 and $active ='A'">
					<!--active A - is Active, V - is inactive -->
					<MessageStatus>0</MessageStatus>
					<rawMessage>
						<xsl:value-of select ="string('RPV')"/>
						<xsl:variable name="timeinSec" select = "($hours * 3600) + ($mins * 60) + $secs"/>
						<xsl:variable name="timePad" select = "substring('00000000',1,5-string-length(string($timeinSec)))"/>
						<xsl:value-of select="concat($timePad,$timeinSec)"/>
						<xsl:choose>
							<xsl:when test = "$eq = 'N'"><xsl:value-of select ="string('+')"/></xsl:when>
							<xsl:otherwise><xsl:value-of select ="string('-')"/></xsl:otherwise>
						</xsl:choose>
<!--CONVERT LATITUDE FOR GLL-->
						<xsl:choose>
							<xsl:when test = "$fromSystem = 'DMS'">
								<xsl:choose>
									<xsl:when test = "$toSystem = 'DD'">
										<xsl:variable name="latValue" select ="(number((number($latM) div 60)) + (number(concat(substring($latm,1,2),'.',substring($latm,3,4)))div 3600))"/>
										<xsl:variable name="latFormat" select ="concat(concat(substring('00',1,2-string-length(string($latD))),$latD), substring-after($latValue,'.'))"/>
										<xsl:value-of select="substring(concat($latFormat,'000000000'),1,7)" />  
									</xsl:when>
								</xsl:choose>
							</xsl:when>
							<xsl:when test = "$fromSystem = 'DDM'">
								<xsl:choose>
									<xsl:when test = "$toSystem = 'DD'">
										<xsl:value-of select="substring(concat($latD,substring-after((number(concat($latM,'.',$latm))div 60),'.'),'000000000'),1,7)" />
									</xsl:when>
								</xsl:choose>
							</xsl:when>
							<xsl:otherwise>
							<!--ALREADY IN FORMAT REQUIRED-->
								<xsl:choose>
									<xsl:when test = "$toSystem = 'DD'">
										<xsl:value-of select="substring(concat(substring-before(concat($lat,'.'),'.'),substring-after(concat($lat,'.'),'.'),'000000000'),1,7)" />
									</xsl:when>
								</xsl:choose>
							</xsl:otherwise>
						</xsl:choose>
<!--END CONVERT LATITUDE FOR GLL-->	
						<xsl:choose>
							<xsl:when test = "$hm = 'E'"><xsl:value-of select ="string('+')"/></xsl:when>
							<xsl:otherwise><xsl:value-of select ="string('-')"/></xsl:otherwise>
						</xsl:choose>
<!--CONVERT LONGITUDE FOR GLL-->
						<xsl:choose>
							<xsl:when test = "$fromSystem = 'DMS'">
								<xsl:choose>
									<xsl:when test = "$toSystem = DD">
										<xsl:variable name="longValue" select ="(number((number($longM) div 60)) + (number(concat(substring($longm,1,2),'.',substring($longm,3,4))))div 3600)"/>
										<xsl:variable name="longFormat" select ="concat(concat(substring('000',1,3-string-length(string($longD))),$longD), substring-after($longValue,'.'))"/>
										<xsl:value-of select="substring(concat($longFormat,'0000000000'),1,8)"/>
									</xsl:when>
								</xsl:choose>
							</xsl:when>
							<xsl:when test = "$fromSystem = 'DDM'">
								<xsl:choose>
									<xsl:when test = "$toSystem = 'DD'">
										<xsl:value-of select="substring(concat($longD,substring-after((number(concat($longM,'.',$longm))div 60),'.'),'0000000000'),1,8)" /> 
									</xsl:when>
								</xsl:choose>
							</xsl:when>
							<xsl:otherwise>
							<!--ALREADY IN FORMAT REQUIRED-->
								<xsl:choose>
									<xsl:when test = "$toSystem = 'DD'">
										<xsl:value-of select="substring(concat(substring-before(concat($long,'.'),'.'),substring-after(concat($long,'.'),'.'),'000000000'),1,8)" />
									</xsl:when>
								</xsl:choose>
							</xsl:otherwise>
						</xsl:choose>
<!--END CONVERT LONGITUDE FOR GLL-->		
						<xsl:value-of select="string('FFFGGG22')"/>
						<xsl:value-of select ="concat(';ID=',$unitID)"/>
						<xsl:value-of select ="concat(';UID=',$aliasID)"/>
					</rawMessage>
					<Type>0</Type>
				</xsl:when>
				<xsl:otherwise>
                                        <MessageStatus>3</MessageStatus>
					<rawMessage>INVALID SATELLITE DATA</rawMessage>
					<Type>0</Type>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:when>
		<!--GPGGA TRANSACTION-->
		<xsl:when test = "$transaction ='GPGGA'" >
			<xsl:variable name="time" select="concat(substring-before(substring-after(.,concat($transaction,',')),','),'00000000')" />
			<xsl:variable name="hours" select="substring($time,1,2)" />
			<xsl:variable name="mins" select="substring($time,3,2)" />
			<xsl:variable name="secs" select="substring($time,5,4)" />
			<xsl:variable name="lat" select="substring-before(substring-after(substring-after(.,concat($transaction,',')),','),',')" />
			<xsl:variable name="latD" select="substring(substring-before($lat,'.'),1,2)" />
			<xsl:variable name="latM" select="substring(substring-before($lat,'.'),3)" />
			<xsl:variable name="latm" select="substring-after($lat,'.')" />
			<xsl:variable name="latMm" select = "concat(concat($latM,'.'),$latm)"/>
			<xsl:variable name="eq" select="substring-before(substring-after(.,concat($lat,',')),',')" />
			<xsl:variable name="long" select="substring-before(substring-after(.,concat($eq,',')),',')" />
			<xsl:variable name="longD" select="substring(substring-before($long,'.'),1,3)" />
			<xsl:variable name="longM" select="substring(substring-before($long,'.'),4)" />
			<xsl:variable name="longm" select="substring-after($long,'.')" />
			<xsl:variable name="longMm" select = "concat(concat($longM,'.'),$longm)"/>
			<xsl:variable name="hm" select="substring-before(substring-after(.,concat($long,',')),',')" />
			<xsl:variable name="qual" select="substring-before(substring-after(.,concat($hm,',')),',')" />
			<xsl:variable name="sats" select="substring-before(substring-after(.,concat($qual,',')),',')" /><xsl:choose>
				<xsl:when test ="string-length($latD)>0 and number($qual)>0">
					<MessageStatus>0</MessageStatus>
					<rawMessage>
						<xsl:value-of select ="string('RPV')"/>
						<xsl:variable name="timeinSec" select = "($hours * 3600) + ($mins * 60) + $secs"/>
						<xsl:variable name="timePad" select = "substring('00000000',1,5-string-length(string($timeinSec)))"/>
						<xsl:value-of select="concat($timePad,$timeinSec)"/>
						<xsl:choose>
							<xsl:when test = "$eq = 'N'"><xsl:value-of select ="string('+')"/></xsl:when>
							<xsl:otherwise><xsl:value-of select ="string('-')"/></xsl:otherwise>
						</xsl:choose>
<!--CONVERT LATITUDE FOR GGA-->								
						<xsl:choose>
							<xsl:when test = "$fromSystem = 'DMS'">
								<xsl:choose>
									<xsl:when test = "$toSystem = 'DD'">
										<xsl:variable name="latValue" select ="(number((number($latM) div 60)) + (number(concat(substring($latm,1,2),'.',substring($latm,3,4)))div 3600))"/>
										<xsl:variable name="latFormat" select ="concat(concat(substring('00',1,2-string-length(string($latD))),$latD), substring-after($latValue,'.'))"/>
										<xsl:value-of select="substring(concat($latFormat,'000000000'),1,7)" />  
									</xsl:when>
								</xsl:choose>
							</xsl:when>
							<xsl:when test = "$fromSystem = 'DDM'">
								<xsl:choose>
									<xsl:when test = "$toSystem = 'DD'">
										<xsl:value-of select="substring(concat($latD,substring-after(number(concat($latM,'.',$latm))div 60,'.'),'000000000'),1,7)" />
									</xsl:when>
								</xsl:choose>
							</xsl:when>
							<xsl:otherwise><!--TREAT ALL ELSE LIKE DD-->
							<!--ALREADY IN FORMAT REQUIRED-->
								<xsl:choose>
									<xsl:when test = "$toSystem = 'DD'">
										<xsl:value-of select="substring(concat(substring-before(concat($lat,'.'),'.'),substring-after(concat($lat,'.'),'.'),'000000000'),1,7)" />
									</xsl:when>
								</xsl:choose>
							</xsl:otherwise>
						</xsl:choose>
<!--END CONVERT LATITUDE FOR GGA-->	
						<xsl:choose>
							<xsl:when test = "$hm = 'E'"><xsl:value-of select ="string('+')"/></xsl:when>
							<xsl:otherwise><xsl:value-of select ="string('-')"/></xsl:otherwise>
						</xsl:choose>
<!--CONVERT LONGITUDE FOR GGA-->
						<xsl:choose>
							<xsl:when test = "$fromSystem = 'DMS'">
								<xsl:choose>
									<xsl:when test = "$toSystem = 'DD'">
										<xsl:variable name="longValue" select ="(number((number($longM) div 60)) + (number(concat(substring($longm,1,2),'.',substring($longm,3,4))))div 3600)"/>
										<xsl:variable name="longFormat" select ="concat(concat(substring('000',1,3-string-length(string($longD))),$longD), substring-after($longValue,'.'))"/>
										<xsl:value-of select="substring(concat($longFormat,'0000000000'),1,8)"/>
									</xsl:when>
								</xsl:choose>
							</xsl:when>
							<xsl:when test = "$fromSystem ='DDM'">
								<xsl:choose>
									<xsl:when test = "$toSystem = 'DD'">
										<xsl:value-of select="substring(concat($longD,substring-after((number(concat($longM,'.',$longm))div 60),'.'),'0000000000'),1,8)" /> 
									</xsl:when>
								</xsl:choose>
							</xsl:when>
							<!--ALREADY IN FORMAT REQUIRED-->
							<xsl:otherwise>
								<xsl:choose>
									<xsl:when test = "$toSystem ='DD'">
										<xsl:value-of select="substring(concat(substring-before(concat($long,'.'),'.'),substring-after(concat($long,'.'),'.'),'000000000'),1,8)" />
									</xsl:when>
								</xsl:choose>
							</xsl:otherwise>
						</xsl:choose>
<!--END CONVERT LONGITUDE FOR GGA-->			
						<xsl:value-of select="string('FFFGGG22')"/>
						<xsl:value-of select ="concat(';ID=',$unitID)"/>
						<xsl:value-of select ="concat(';UID=',$aliasID)"/>
					</rawMessage>
					<Type>0</Type>
				</xsl:when>
				<xsl:otherwise>
					<MessageStatus>3</MessageStatus>
					<rawMessage>INVALID SATELLITE DATA</rawMessage>
					<Type>0</Type>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:when>
		<!--GPRMC TRANSACTION-->
		<xsl:when test = "$transaction ='GPRMC'">
			<xsl:variable name="time" select="concat(substring-before(substring-after(.,concat($transaction,',')),','),'00000000')" />
			<xsl:variable name="hours" select="substring($time,1,2)" />
			<xsl:variable name="mins" select="substring($time,3,2)" />
			<xsl:variable name="secs" select="substring($time,5,2)" />
			<xsl:variable name="status" select="substring-before(substring-after(substring-after(.,concat($transaction,',')),','),',')" />
			<xsl:variable name="lat" select="substring-before(substring-after(.,concat($status,',')),',')" />
			<xsl:variable name="latD" select="substring(substring-before($lat,'.'),1,2)" />
			<xsl:variable name="latM" select="substring(substring-before($lat,'.'),3)" />
			<xsl:variable name="latm" select="substring-after($lat,'.')" />
			<xsl:variable name="latMm" select = "concat(concat($latM,'.'),$latm)"/>
			<xsl:variable name="eq" select="substring-before(substring-after(.,concat($lat,',')),',')" />
			<xsl:variable name="long" select="substring-before(substring-after(.,concat($eq,',')),',')" />
			<xsl:variable name="longD" select="substring(substring-before($long,'.'),1,3)" />
			<xsl:variable name="longM" select="substring(substring-before($long,'.'),4)" />
			<xsl:variable name="longm" select="substring-after($long,'.')" />
			<xsl:variable name="longMm" select = "concat(concat($longM,'.'),$longm)"/>
			<xsl:variable name="hm" select="substring-before(substring-after(.,concat($long,',')),',')" />
			<xsl:choose>
			<xsl:when test ="string-length($latD)>0 and $status ='A'">
				<MessageStatus>0</MessageStatus>
				<rawMessage>
					<xsl:value-of select ="string('RPV')"/>
					<xsl:variable name="timeinSec" select = "($hours * 3600) + ($mins * 60) + $secs"/>
					<xsl:variable name="timePad" select = "substring('00000000',1,5-string-length(string($timeinSec)))"/>
					<xsl:value-of select="concat($timePad,$timeinSec)"/>
					<xsl:choose>
						<xsl:when test = "$eq = 'N'"><xsl:value-of select ="string('+')"/></xsl:when>
						<xsl:otherwise><xsl:value-of select ="string('-')"/></xsl:otherwise>
					</xsl:choose>
<!--CONVERT LATITUDE FOR RMC-->
					<xsl:choose>
						<xsl:when test = "$fromSystem = 'DMS'">
							<xsl:choose>
								<xsl:when test = "$toSystem = 'DD'">
									<xsl:variable name="latValue" select ="(number((number($latM) div 60)) + (number(concat(substring($latm,1,2),'.',substring($latm,3,4)))div 3600))"/>
									<xsl:variable name="latFormat" select ="concat(concat(substring('00',1,2-string-length(string($latD))),$latD), substring-after($latValue,'.'))"/>
									<xsl:value-of select="substring(concat($latFormat,'000000000'),1,7)" />  
								</xsl:when>
							</xsl:choose>
						</xsl:when>
						<xsl:when test = "$fromSystem = 'DDM'">
							<xsl:choose>
								<xsl:when test = "$toSystem = 'DD'">
									<xsl:value-of select="substring(concat($latD,substring-after((number(concat($latM,'.',$latm))div 60),'.'),'000000000'),1,7)" />
								</xsl:when>
							</xsl:choose>
						</xsl:when>
						<!--ALREADY IN FORMAT REQUIRED-->
						<xsl:otherwise>
							<xsl:choose>
								<xsl:when test = "$toSystem = 'DD'">
									<xsl:value-of select="substring(concat(substring-before(concat($lat,'.'),'.'),substring-after(concat($lat,'.'),'.'),'000000000'),1,7)" />
								</xsl:when>
							</xsl:choose>
						</xsl:otherwise>
					</xsl:choose>
<!--END CONVERT LATITUDE FOR RMC-->							
					<xsl:choose>
						<xsl:when test = "$hm = 'E'"><xsl:value-of select ="string('+')"/></xsl:when>
						<xsl:otherwise><xsl:value-of select ="string('-')"/></xsl:otherwise>
					</xsl:choose>
<!--CONVERT LONGITUDE FOR RMC-->
					<xsl:choose>
						<xsl:when test = "$fromSystem = 'DMS'">
							<xsl:choose>
								<xsl:when test = "$toSystem = 'DD'">
									<xsl:variable name="longValue" select ="(number((number($longM) div 60)) + (number(concat(substring($longm,1,2),'.',substring($longm,3,4))))div 3600)"/>
									<xsl:variable name="longFormat" select ="concat(concat(substring('000',1,3-string-length(string($longD))),$longD), substring-after($longValue,'.'))"/>
									<xsl:value-of select="substring(concat($longFormat,'0000000000'),1,8)"/>
								</xsl:when>
							</xsl:choose>
						</xsl:when>
						<xsl:when test = "$fromSystem = 'DDM'">
							<xsl:choose>
								<xsl:when test = "$toSystem = 'DD'">
									<xsl:value-of select="substring(concat($longD,substring-after((number(concat($longM,'.',$longm))div 60),'.'),'0000000000'),1,8)" /> 
								</xsl:when>
							</xsl:choose>
						</xsl:when>
						<!--ALREADY IN FORMAT REQUIRED-->
						<xsl:otherwise>
							<xsl:choose>
								<xsl:when test = "$toSystem = 'DD'">
									<xsl:value-of select="substring(concat(substring-before(concat($long,'.'),'.'),substring-after(concat($long,'.'),'.'),'000000000'),1,8)" />
								</xsl:when>
							</xsl:choose>
						</xsl:otherwise>
					</xsl:choose>
<!--END CONVERT LONGITUDE FOR RMC-->			
					<xsl:value-of select="string('FFFGGG22')"/>
					<xsl:value-of select ="concat(';ID=',$unitID)"/>
					<xsl:value-of select ="concat(';UID=',$aliasID)"/>
				</rawMessage>
				<Type>0</Type>
			</xsl:when>
			<xsl:otherwise>
				<MessageStatus>3</MessageStatus>
				<rawMessage>INVALID SATELLITE DATA</rawMessage>
				<Type>0</Type>
			</xsl:otherwise>
			</xsl:choose>
		</xsl:when>
		<xsl:otherwise>
			<MessageStatus>3</MessageStatus>
			<rawMessage>NOT A VALID TRANSACTION TYPE</rawMessage>
			<Type>1</Type>
		</xsl:otherwise>
		</xsl:choose>
		</GPSTransform>
	</xsl:template>
</xsl:stylesheet>