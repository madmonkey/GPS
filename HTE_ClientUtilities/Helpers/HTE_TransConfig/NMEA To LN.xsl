<xsl:stylesheet xmlns:xsl = "http://www.w3.org/1999/XSL/Transform" version = "1.0" > 
<xsl:output method = "xml" standalone ="yes" encoding="UTF-8"/> 
	<xsl:template match="/">
		<GPSTransform>
		<xsl:variable name="transaction" select="substring-before(.,',')"/>
		<xsl:variable name = "toDecimal" select = "'false'"/>
		<xsl:variable name = "toSeconds" select = "'true'"/>
		<xsl:choose>
		<!--Lat/Lon data - earlier G-12s do NOT transmit -->
		<!--Geographic Latitude and Longitude, holdover from Lorain data, may be prefixed with LC ie GP-->
		
		<!--GPGLL TRANSACTION-->
		<xsl:when test = "$transaction ='GPGLL'" >
			<!--FORMATTED TO TAIP LN TRANSACTION-->
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
			<xsl:variable name="time" select="substring-before(substring(substring-after(.,$hm),2),',')" />
			<xsl:variable name="hours" select="substring($time,1,2)" />
			<xsl:variable name="mins" select="substring($time,3,2)" />
			<xsl:variable name="secs" select="substring($time,5,2)" />
			<xsl:variable name="active" select="substring-before(substring(substring-after(.,$time),2),',')" />
			<xsl:variable name="chksum" select="substring(substring-after(.,$active),2)" />
			<!--0 - processed, 1 - processed warn, 2 - processed error, 3 - error -->
					<xsl:choose>
					<xsl:when test ="string-length($latD)>0">
						<MessageStatus>0</MessageStatus>
						<rawMessage>
							<xsl:value-of select ="string('RLN')"/>
							<!--since a fixed format, make sure spacing is in accordance with specification !-->
							<xsl:variable name="timeinSec" select = "($hours * 3600) + ($mins * 60) + $secs"/>
							<xsl:variable name="timePad" select = "substring('00000000',1,5-string-length(string($timeinSec)))"/>
							<xsl:value-of select="concat($timePad,$timeinSec,'000')"/>
							<xsl:choose>
								<xsl:when test = "$eq = 'N'"><xsl:value-of select ="string('+')"/></xsl:when>
								<xsl:otherwise><xsl:value-of select ="string('-')"/></xsl:otherwise>
							</xsl:choose>
							<xsl:choose>
							<xsl:when test = "$toDecimal = 'true'">
								<xsl:variable name="latValue" select ="(number((number($latM) div 60)) + (number(concat(substring($latm,1,2),'.',substring($latm,3,4)))div 3600))"/>
								<xsl:variable name="latFormat" select ="concat(concat(substring('00',1,2-string-length(string($latD))),$latD), substring-after($latValue,'.'))"/>
								<xsl:value-of select="substring(concat($latFormat,'000000000'),1,9)" />  
							</xsl:when>
							<xsl:otherwise>
								<xsl:if test = "$toSeconds= 'true'">
									<xsl:value-of select="substring(concat($latD,$latM,($latm * 60),'000000000'),1,9)" />
								</xsl:if>
								<xsl:if test = "$toSeconds= 'false'">
									<xsl:value-of select="substring(concat($latD,$latM,$latm,'000000000'),1,9)" />
								</xsl:if>
							</xsl:otherwise>
							</xsl:choose>
							<xsl:choose>
								<xsl:when test = "$hm = 'E'"><xsl:value-of select ="string('+')"/></xsl:when>
								<xsl:otherwise><xsl:value-of select ="string('-')"/></xsl:otherwise>
							</xsl:choose>
							<xsl:choose>
							<xsl:when test = "$toDecimal = 'true'">
								<xsl:variable name="longValue" select ="(number((number($longM) div 60)) + (number(concat(substring($longm,1,2),'.',substring($longm,3,4))))div 3600)"/>
								<xsl:variable name="longFormat" select ="concat(concat(substring('000',1,3-string-length(string($longD))),$longD), substring-after($longValue,'.'))"/>
								<xsl:value-of select="substring(concat($longFormat,'0000000000'),1,10)"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:if test = "$toSeconds= 'true'">
									<xsl:value-of select="substring(concat($longD,$longM,($longm*60),'0000000000'),1,10)" /> 
								</xsl:if> 
								<xsl:if test = "$toSeconds= 'false'">
									<xsl:value-of select="substring(concat($longD,$longM,$longm,'0000000000'),1,10)" /> 
								</xsl:if> 
							</xsl:otherwise>
							</xsl:choose>		
							<xsl:value-of select="string('+GGGGGGHHIIIJ+KKKLMMMN01PPQQRRRRRRRRRR90')"/>
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
		
			<!--GPGGA TRANSACTION-->
		<xsl:when test = "$transaction ='GPGGA'">
			<!--0 - processed, 1 - processed warn, 2 - processed error, 3 - error -->
			<xsl:variable name="time" select="substring-before(substring-after(.,concat($transaction,',')),',')" />
			<xsl:variable name="hours" select="substring($time,1,2)" />
			<xsl:variable name="mins" select="substring($time,3,2)" />
			<xsl:variable name="secs" select="substring($time,5,4)" />
			<xsl:variable name="lat" select="substring-before(substring-after(.,concat($time,',')),',')" />
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
			<xsl:variable name="sats" select="substring-before(substring-after(.,concat($qual,',')),',')" />
			<!--0 - processed, 1 - processed warn, 2 - processed error, 3 - error -->
					<xsl:choose>
					<xsl:when test ="string-length($latD)>0">
						<MessageStatus>0</MessageStatus>
						<rawMessage>
							<xsl:value-of select ="string('RLN')"/>
							<!--since a fixed format, make sure spacing is in accordance with specification !-->
							<xsl:variable name="timeinSec" select = "($hours * 3600) + ($mins * 60) + $secs"/>
							<xsl:variable name="timePad" select = "substring('00000000',1,5-string-length(string($timeinSec)))"/>
							<xsl:value-of select="concat($timePad,$timeinSec,'000')"/>
							<xsl:choose>
								<xsl:when test = "$eq = 'N'"><xsl:value-of select ="string('+')"/></xsl:when>
								<xsl:otherwise><xsl:value-of select ="string('-')"/></xsl:otherwise>
							</xsl:choose>
							<xsl:choose>
							<xsl:when test = "$toDecimal = 'true'">
								<xsl:variable name="latValue" select ="(number((number($latM) div 60)) + (number(concat(substring($latm,1,2),'.',substring($latm,3,4)))div 3600))"/>
								<xsl:variable name="latFormat" select ="concat(concat(substring('00',1,2-string-length(string($latD))),$latD), substring-after($latValue,'.'))"/>
								<xsl:value-of select="substring(concat($latFormat,'000000000'),1,9)" />  
							</xsl:when>
							<xsl:otherwise>
								<xsl:if test = "$toSeconds= 'true'">
									<xsl:value-of select="substring(concat($latD,$latM,($latm * 60),'000000000'),1,9)" />
								</xsl:if>
								<xsl:if test = "$toSeconds= 'false'">
									<xsl:value-of select="substring(concat($latD,$latM,$latm,'000000000'),1,9)" />
								</xsl:if>
							</xsl:otherwise>
							</xsl:choose>
							<xsl:choose>
								<xsl:when test = "$hm = 'E'"><xsl:value-of select ="string('+')"/></xsl:when>
								<xsl:otherwise><xsl:value-of select ="string('-')"/></xsl:otherwise>
							</xsl:choose>
							<xsl:choose>
							<xsl:when test = "$toDecimal = 'true'">
								<xsl:variable name="longValue" select ="(number((number($longM) div 60)) + (number(concat(substring($longm,1,2),'.',substring($longm,3,4))))div 3600)"/>
								<xsl:variable name="longFormat" select ="concat(concat(substring('000',1,3-string-length(string($longD))),$longD), substring-after($longValue,'.'))"/>
								<xsl:value-of select="substring(concat($longFormat,'0000000000'),1,10)"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:if test = "$toSeconds= 'true'">
									<xsl:value-of select="substring(concat($longD,$longM,($longm*60),'0000000000'),1,10)" /> 
								</xsl:if> 
								<xsl:if test = "$toSeconds= 'false'">
									<xsl:value-of select="substring(concat($longD,$longM,$longm,'0000000000'),1,10)" /> 
								</xsl:if> 
							</xsl:otherwise>
							</xsl:choose>		
							<xsl:value-of select="string('+GGGGGGHHIIIJ+KKKLMMMN01PPQQRRRRRRRRRR90')"/>
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
			
			<!--GPRMC TRANSACTION-->
		<xsl:when test = "$transaction ='GPRMC'">
			<xsl:variable name="time" select="substring-before(substring-after(.,concat($transaction,',')),',')" />
			<xsl:variable name="hours" select="substring($time,1,2)" />
			<xsl:variable name="mins" select="substring($time,3,2)" />
			<xsl:variable name="secs" select="substring($time,5,4)" />
			<xsl:variable name="status" select="substring-before(substring-after(.,concat($time,',')),',')" />
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
			<!--0 - processed, 1 - processed warn, 2 - processed error, 3 - error -->
					<xsl:choose>
					<xsl:when test ="string-length($latD)>0">
						<MessageStatus>0</MessageStatus>
						<rawMessage>
							<xsl:value-of select ="string('RLN')"/>
							<!--since a fixed format, make sure spacing is in accordance with specification !-->
							<xsl:variable name="timeinSec" select = "($hours * 3600) + ($mins * 60) + $secs"/>
							<xsl:variable name="timePad" select = "substring('00000000',1,5-string-length(string($timeinSec)))"/>
							<xsl:value-of select="concat($timePad,$timeinSec,'000')"/>
							<xsl:choose>
								<xsl:when test = "$eq = 'N'"><xsl:value-of select ="string('+')"/></xsl:when>
								<xsl:otherwise><xsl:value-of select ="string('-')"/></xsl:otherwise>
							</xsl:choose>
							<xsl:choose>
							<xsl:when test = "$toDecimal = 'true'">
								<xsl:variable name="latValue" select ="(number((number($latM) div 60)) + (number(concat(substring($latm,1,2),'.',substring($latm,3,4)))div 3600))"/>
								<xsl:variable name="latFormat" select ="concat(concat(substring('00',1,2-string-length(string($latD))),$latD), substring-after($latValue,'.'))"/>
								<xsl:value-of select="substring(concat($latFormat,'000000000'),1,9)" />  
							</xsl:when>
							<xsl:otherwise>
								<xsl:if test = "$toSeconds= 'true'">
									<xsl:value-of select="substring(concat($latD,$latM,($latm * 60),'000000000'),1,9)" />
								</xsl:if>
								<xsl:if test = "$toSeconds= 'false'">
									<xsl:value-of select="substring(concat($latD,$latM,$latm,'000000000'),1,9)" />
								</xsl:if>
							</xsl:otherwise>
							</xsl:choose>
							<xsl:choose>
								<xsl:when test = "$hm = 'E'"><xsl:value-of select ="string('+')"/></xsl:when>
								<xsl:otherwise><xsl:value-of select ="string('-')"/></xsl:otherwise>
							</xsl:choose>
							<xsl:choose>
							<xsl:when test = "$toDecimal = 'true'">
								<xsl:variable name="longValue" select ="(number((number($longM) div 60)) + (number(concat(substring($longm,1,2),'.',substring($longm,3,4))))div 3600)"/>
								<xsl:variable name="longFormat" select ="concat(concat(substring('000',1,3-string-length(string($longD))),$longD), substring-after($longValue,'.'))"/>
								<xsl:value-of select="substring(concat($longFormat,'0000000000'),1,10)"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:if test = "$toSeconds= 'true'">
									<xsl:value-of select="substring(concat($longD,$longM,($longm*60),'0000000000'),1,10)" /> 
								</xsl:if> 
								<xsl:if test = "$toSeconds= 'false'">
									<xsl:value-of select="substring(concat($longD,$longM,$longm,'0000000000'),1,10)" /> 
								</xsl:if> 
							</xsl:otherwise>
							</xsl:choose>		
							<xsl:value-of select="string('+GGGGGGHHIIIJ+KKKLMMMN01PPQQRRRRRRRRRR90')"/>
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
			<Type>1</Type>
		</xsl:otherwise>
		</xsl:choose>
		</GPSTransform>
	</xsl:template>
</xsl:stylesheet>