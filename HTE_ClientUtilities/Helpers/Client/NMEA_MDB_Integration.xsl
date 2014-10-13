<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" standalone="yes" encoding="UTF-8" version="2.0" />
	<!--NMEA_MDB_Integration.xsl (c) SunGard Public Sector Inc-->
	<!--NMEA_MDB_Integration STYLESHEET USED FOR FORMATTING GPS INFORMATION FOR THE MOBILE DATA BROWSER (c)-->
        <!--UNLIKE MOST THE DECIMAL POINT IS NOT IMPLIED AND THIS IS NOT A FIXED_SIZE FORMAT-->
	<!--INCLUDE THE COORDINATE SYSTEM FROM YOUR INCOMING DATA DEVICE 
		CHOICES ARE: 
		'DMS' (Degrees Minutes Seconds), 
		'DDM' (Degrees Decimal Minutes) NMEA 0183 Standard,
		'DD'  (Decimal Degrees)
            -->
	<xsl:variable name="fromSystem" select="'DDM'" />
	<!--WHAT COORDINATE SYSTEM DOES THE OUTPUT NEED TO BE 
		CHOICES ARE:
		'DD' (Decimal Degrees)
            -->
	<xsl:variable name="toSystem" select="'DD'" />
	<xsl:template match="/">
		<GPSTransform>
			<xsl:variable name="transaction" select="substring-before(.,',')" />
			<xsl:choose>
				<!--GPGLL TRANSACTION-->
				<xsl:when test="$transaction ='GPGLL' or $transaction ='$GPGLL'">
					<!--GPGLL,4916.45,N,12311.12,W,225444,A -->
					<xsl:variable name="lat" select="substring-before(substring-after(.,concat($transaction,',')),',')" />
					<xsl:variable name="latD" select="substring(substring-before($lat,'.'),1,2)" />
					<xsl:variable name="latM" select="substring(substring-before($lat,'.'),3)" />
					<xsl:variable name="latm" select="substring-after($lat,'.')" />
					<xsl:variable name="latMm" select="concat(concat($latM,'.'),$latm)" />
					<xsl:variable name="eq" select="substring(substring-after(.,$lat),2,1)" />
					<xsl:variable name="long" select="substring-before(substring(substring-after(.,$eq),2),',')" />
					<xsl:variable name="longD" select="substring(substring-before($long,'.'),1,3)" />
					<xsl:variable name="longM" select="substring(substring-before($long,'.'),4)" />
					<xsl:variable name="longm" select="substring-after($long,'.')" />
					<xsl:variable name="longMm" select="concat(concat($longM,'.'),$longm)" />
					<xsl:variable name="hm" select="substring-before(substring(substring-after(.,$long),2),',')" />
					<xsl:variable name="time" select="substring-before(substring(substring-after(.,$hm),2),',')" />
					<xsl:variable name="hours" select="substring($time,1,2)" />
					<xsl:variable name="mins" select="substring($time,3,2)" />
					<xsl:variable name="secs" select="substring($time,5,2)" />
					<xsl:variable name="active" select="substring(substring-after(.,$time),2,1)" />
					<xsl:variable name="chksum" select="substring(substring-after(.,$active),2,2)" />
					<!-- Potential Error Codes and Severity Level(s) -->
					<!--0 - processed, 1 - processed warn, 2 - processed error, 3 - error -->
					<xsl:choose>
						<xsl:when test="$active='A'">
							<!-- Data is valid, go ahead and process -->
							<MessageStatus>0</MessageStatus>
							<Type>1</Type>
							<rawMessage method="xml">
								<MDBGPS>
									<LAT>
										<xsl:choose>
											<xsl:when test="$eq = 'N'">
												<xsl:value-of select="string('+')" />
											</xsl:when>
											<xsl:otherwise>
												<xsl:value-of select="string('-')" />
											</xsl:otherwise>
										</xsl:choose>
										<xsl:choose>
											<xsl:when test="$fromSystem = 'DMS'">
												<xsl:choose>
													<xsl:when test="$toSystem = 'DD'">
														<xsl:variable name="latValue" select="(number(number($latM) div 60) + (number(concat(substring($latm,1,2), '.', substring($latm,3,4))) div 3600))" />
														<xsl:variable name="latFormat" select="concat($latD,'.',substring-after($latValue,'.'))"/>
														<xsl:value-of select="$latFormat" />
													</xsl:when>
												</xsl:choose>
											</xsl:when>
											<xsl:when test="$fromSystem = 'DDM'">
												<xsl:choose>
													<xsl:when test="$toSystem = 'DD'">
														<xsl:value-of select="concat($latD, '.', substring-after(number(concat($latM,'.',$latm)div 60),'.'))" />
													</xsl:when>
												</xsl:choose>
											</xsl:when>
											<!--ALREADY DD - GOOD FOR US-->
                                                                                        <xsl:otherwise>
												<!--ALREADY IN FORMAT REQUIRED-->
												<xsl:choose>
													<xsl:when test="$toSystem = 'DD'">
														<xsl:value-of select="concat($latD,'.', $latM, $latm)" />
													</xsl:when>
												</xsl:choose>
											</xsl:otherwise>
										</xsl:choose>
									</LAT>
									<LONG>
										<!--CONVERT LONGITUDE FOR GLL-->
										<xsl:choose>
											<xsl:when test="$hm = 'E'">
												<xsl:value-of select="string('+')" />
											</xsl:when>
											<xsl:otherwise>
												<xsl:value-of select="string('-')" />
											</xsl:otherwise>
										</xsl:choose>
										<xsl:choose>
											<xsl:when test="$fromSystem = 'DMS'">
												<xsl:choose>
													<xsl:when test="$toSystem = 'DD'">
														<xsl:variable name="longValue" select="(number((number($longM) div 60)) + (number(concat(substring($longm,1,2),'.',substring($longm,3,4))))div 3600)" />
														<xsl:variable name="longFormat" select="concat($longD, '.', substring-after($longValue,'.'))" />
														<xsl:value-of select="$longFormat" />
													</xsl:when>
												</xsl:choose>
											</xsl:when>
											<xsl:when test="$fromSystem = 'DDM'">
												<xsl:choose>
													<xsl:when test="$toSystem = 'DD'">
														<xsl:value-of select="concat($longD,'.', substring-after((number(concat($longM,'.',$longm))div 60),'.'))" />
													</xsl:when>
												</xsl:choose>
											</xsl:when>
											<xsl:otherwise>
												<!--ALREADY IN FORMAT REQUIRED-->
												<xsl:choose>
													<xsl:when test="$toSystem = 'DD'">
														<xsl:value-of select="concat($longD,'.',substring-before($longMm,'.'),substring-after($longMm,'.'))" />
													</xsl:when>
												</xsl:choose>
											</xsl:otherwise>
										</xsl:choose>
										<!-- END CONVERT LONGITUDE FOR GLL -->
									</LONG>
									<TIMESTAMP>
                                                                            <xsl:value-of select="concat($hours,':', $mins, ':', $secs)" />
                                                                                <!-- IF WE WANTED SECONDS FROM MIDNIGHT - LIKE TAIP THEN
										<xsl:variable name="timeinSec" select="($hours * 3600) + ($mins * 60) + $secs" />
										<xsl:value-of select="$timeinSec" />
                                                                                -->
									</TIMESTAMP>
								</MDBGPS>
							</rawMessage>
						</xsl:when>
						<xsl:otherwise>
							<!-- Data is invalid, go ahead and discard -->
							<MessageStatus>3</MessageStatus>
							<Type>1</Type>
							<rawMessage>Active Flag [<xsl:value-of select="$active" />] is disabled!</rawMessage>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<!--GPGGA TRANSACTION-->
				<xsl:when test="$transaction ='GPGGA' or $transaction ='$GPGGA'">
					<xsl:variable name="time" select="concat(substring-before(substring-after(.,concat($transaction,',')),','),'00000000')" />
					<xsl:variable name="hours" select="substring($time,1,2)" />
					<xsl:variable name="mins" select="substring($time,3,2)" />
					<xsl:variable name="secs" select="substring($time,5,2)" />
					<xsl:variable name="lat" select="substring-before(substring-after(substring-after(.,concat($transaction,',')),','),',')" />
					<xsl:variable name="latD" select="substring(substring-before($lat,'.'),1,2)" />
					<xsl:variable name="latM" select="substring(substring-before($lat,'.'),3)" />
					<xsl:variable name="latm" select="substring-after($lat,'.')" />
					<xsl:variable name="latMm" select="concat(concat($latM,'.'),$latm)" />
					<xsl:variable name="eq" select="substring-before(substring-after(.,concat($lat,',')),',')" />
					<xsl:variable name="long" select="substring-before(substring-after(.,concat($eq,',')),',')" />
					<xsl:variable name="longD" select="substring(substring-before($long,'.'),1,3)" />
					<xsl:variable name="longM" select="substring(substring-before($long,'.'),4)" />
					<xsl:variable name="longm" select="substring-after($long,'.')" />
					<xsl:variable name="longMm" select="concat(concat($longM,'.'),$longm)" />
					<xsl:variable name="hm" select="substring-before(substring-after(.,concat($long,',')),',')" />
					<xsl:variable name="qual" select="substring-before(substring-after(.,concat($hm,',')),',')" />
					<xsl:variable name="sats" select="substring-before(substring-after(.,concat($qual,',')),',')" />
					<xsl:choose>
						<xsl:when test="$qual &gt;'0'">
							<MessageStatus>0</MessageStatus>
							<Type>1</Type>
							<rawMessage method="xml">
								<MDBGPS>
									<LAT>
									    <!--CONVERT LATITUDE FOR GGA-->
                                                                            <xsl:choose>
                                                                                <xsl:when test="$eq = 'N'">
                                                                                    <xsl:value-of select="string('+')" />
                                                                                </xsl:when>
                                                                                <xsl:otherwise>
                                                                                    <xsl:value-of select="string('-')" />
                                                                                </xsl:otherwise>
                                                                            </xsl:choose>
                                                                            <xsl:choose>
										<xsl:when test="$fromSystem = 'DMS'">
                                                                                    <xsl:choose>
                                                                                        <xsl:when test="$toSystem = 'DD'">
                                                                                            <xsl:variable name="latValue" select="(number((number($latM) div 60)) + (number(concat(substring($latm,1,2),'.',substring($latm,3,4)))div 3600))" />
                                                                                            <xsl:variable name="latFormat" select="concat($latD, '.', substring-after($latValue,'.'))" />
                                                                                            <xsl:value-of select="$latFormat" />
											</xsl:when>
                                                                                    </xsl:choose>
										</xsl:when>
										<xsl:when test="$fromSystem = 'DDM'">
                                                                                    <xsl:choose>
                                                                                        <xsl:when test="$toSystem = 'DD'">
                                                                                            <xsl:value-of select="concat($latD, '.', substring-after(number(concat($latM,'.',$latm))div 60,'.'))" />
                                                                                        </xsl:when>
                                                                                    </xsl:choose>
										</xsl:when>
										<xsl:otherwise>
                                                                                    <!--ALREADY IN FORMAT REQUIRED-->
                                                                                    <xsl:choose>
                                                                                        <xsl:when test="$toSystem = 'DD'">
                                                                                            <xsl:value-of select="concat($latD,'.',substring-before($latMm,'.'), substring-after($latMm,'.'))" />
                                                                                        </xsl:when>
                                                                                    </xsl:choose>
										</xsl:otherwise>
									    </xsl:choose>
										<!--END CONVERT LATITUDE FOR GGA-->
									</LAT>
									<LONG>
									    <!--CONVERT LONGITUDE FOR GGA-->
                                                                            <xsl:choose>
                                                                                <xsl:when test="$hm = 'E'">
                                                                                    <xsl:value-of select="string('+')" />
										</xsl:when>
										<xsl:otherwise>
                                                                                    <xsl:value-of select="string('-')" />
										</xsl:otherwise>
									    </xsl:choose>
                                                                            <xsl:choose>
										<xsl:when test="$fromSystem = 'DMS'">
                                                                                    <xsl:choose>
                                                                                        <xsl:when test="$toSystem = 'DD'">
                                                                                            <xsl:variable name="longValue" select="(number((number($longM) div 60)) + (number(concat(substring($longm,1,2),'.',substring($longm,3,4))))div 3600)" />
                                                                                            <xsl:variable name="longFormat" select="concat($longD,'.', substring-after($longValue,'.'))" />
                                                                                            <xsl:value-of select="$longFormat" />
                                                                                        </xsl:when>
                                                                                    </xsl:choose>
										</xsl:when>
										<xsl:when test="$fromSystem ='DDM'">
                                                                                    <xsl:choose>
                                                                                        <xsl:when test="$toSystem = 'DD'">
                                                                                            <xsl:value-of select="concat($longD,'.',substring-after((number(concat($longM,'.',$longm))div 60),'.'))" />
                                                                                        </xsl:when>
                                                                                    </xsl:choose>
										</xsl:when>
										<!--ALREADY IN FORMAT REQUIRED-->
										<xsl:otherwise>
                                                                                    <xsl:choose>
                                                                                        <xsl:when test="$toSystem ='DD'">
                                                                                            <xsl:value-of select="concat($longD,'.',substring-before($longMm,'.'),substring-after($longMm,'.'))" />
                                                                                        </xsl:when>
                                                                                    </xsl:choose>
										</xsl:otherwise>
									    </xsl:choose>
										<!--END CONVERT LONGITUDE FOR GGA-->
									</LONG>
									<TIMESTAMP>
										<xsl:value-of select="concat($hours,':', $mins, ':', $secs)" />
                                                                                <!-- IF WE WANTED SECONDS FROM MIDNIGHT - LIKE TAIP THEN
                                                                                <xsl:variable name="timeinSec" select="($hours * 3600) + ($mins * 60) + $secs" />
										<xsl:value-of select="$timeinSec"/>
                                                                                -->
									</TIMESTAMP>
								</MDBGPS>
							</rawMessage>
						</xsl:when>
						<xsl:otherwise>
							<!-- Data is invalid, go ahead and discard -->
							<MessageStatus>3</MessageStatus>
							<Type>1</Type>
							<rawMessage>Quality of data is poor!</rawMessage>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<!-- GPRMC TRANSACTION -->
				<xsl:when test="$transaction ='GPRMC' or $transaction ='$GPRMC'">
					<xsl:variable name="time" select="concat(substring-before(substring-after(.,concat($transaction,',')),','),'00000000')" />
					<xsl:variable name="hours" select="substring($time,1,2)" />
					<xsl:variable name="mins" select="substring($time,3,2)" />
					<xsl:variable name="secs" select="substring($time,5,2)" />
					<xsl:variable name="status" select="substring-before(substring-after(substring-after(.,concat($transaction,',')),','),',')" />
					<xsl:variable name="lat" select="substring-before(substring-after(.,concat($status,',')),',')" />
					<xsl:variable name="latD" select="substring(substring-before($lat,'.'),1,2)" />
					<xsl:variable name="latM" select="substring(substring-before($lat,'.'),3)" />
					<xsl:variable name="latm" select="substring-after($lat,'.')" />
					<xsl:variable name="latMm" select="concat(concat($latM,'.'),$latm)" />
					<xsl:variable name="eq" select="substring-before(substring-after(.,concat($lat,',')),',')" />
					<xsl:variable name="long" select="substring-before(substring-after(.,concat($eq,',')),',')" />
					<xsl:variable name="longD" select="substring(substring-before($long,'.'),1,3)" />
					<xsl:variable name="longM" select="substring(substring-before($long,'.'),4)" />
					<xsl:variable name="longm" select="substring-after($long,'.')" />
					<xsl:variable name="longMm" select="concat(concat($longM,'.'),$longm)" />
					<xsl:variable name="hm" select="substring-before(substring-after(.,concat($long,',')),',')" />
					<xsl:choose>
						<xsl:when test="$status='A'">
						    <MessageStatus>0</MessageStatus>
                                                    <Type>1</Type>
                                                    <rawMessage method="xml">
                                                        <MDBGPS>
                                                            <LAT>
                                                                <!--CONVERT LATITUDE FOR RMC-->
								<xsl:choose>
                                                                    <xsl:when test="$eq = 'N'">
									<xsl:value-of select="string('+')" />
                                                                    </xsl:when>
                                                                    <xsl:otherwise>
                                                                        <xsl:value-of select="string('-')" />
                                                                    </xsl:otherwise>
								</xsl:choose>
								<xsl:choose>
                                                                    <xsl:when test="$fromSystem = 'DMS'">
                                                                        <xsl:choose>
                                                                            <xsl:when test="$toSystem = 'DD'">
										<xsl:variable name="latValue" select="(number((number($latM) div 60)) + (number(concat(substring($latm,1,2),'.',substring($latm,3,4)))div 3600))" />
										<xsl:variable name="latFormat" select="concat($latD, '.', substring-after($latValue,'.'))" />
                                                                                <xsl:value-of select="$latFormat" />
                                                                            </xsl:when>
									</xsl:choose>
								    </xsl:when>
                                                                    <xsl:when test="$fromSystem = 'DDM'">
                                                                        <xsl:choose>
                                                                            <xsl:when test="$toSystem = 'DD'">
                                                                                <xsl:value-of select="concat($latD, '.', substring-after(number(concat($latM,'.',$latm))div 60,'.'))" />
                                                                            </xsl:when>
                                                                        </xsl:choose>
                                                                    </xsl:when>
                                                                    <!--ALREADY IN FORMAT REQUIRED-->
                                                                    <xsl:otherwise>
                                                                        <xsl:choose>
                                                                            <xsl:when test="$toSystem = 'DD'">
                                                                                <xsl:value-of select="concat($latD,'.',substring-before($latMm,'.'), substring-after($latMm,'.'))" />
                                                                            </xsl:when>
                                                                        </xsl:choose>
                                                                    </xsl:otherwise>
								</xsl:choose>
								<!--END CONVERT LATITUDE FOR RMC-->
							    </LAT>
							    <LONG>
                                                                <!--CONVERT LONGITUDE FOR RMC-->
                                                                <xsl:choose>
                                                                    <xsl:when test="$hm = 'E'">
                                                                        <xsl:value-of select="string('+')" />
                                                                    </xsl:when>
                                                                    <xsl:otherwise>
                                                                        <xsl:value-of select="string('-')" />
                                                                    </xsl:otherwise>
                                                                </xsl:choose>
                                                                <xsl:choose>
                                                                    <xsl:when test="$fromSystem = 'DMS'">
                                                                        <xsl:choose>
                                                                            <xsl:when test="$toSystem = 'DD'">
                                                                                <xsl:variable name="longValue" select="(number((number($longM) div 60)) + (number(concat(substring($longm,1,2),'.',substring($longm,3,4))))div 3600)" />
                                                                                <xsl:variable name="longFormat" select="concat($longD,'.', substring-after($longValue,'.'))" />
                                                                                <xsl:value-of select="$longFormat" />
                                                                            </xsl:when>
                                                                        </xsl:choose>
                                                                    </xsl:when>
                                                                    <xsl:when test="$fromSystem = 'DDM'">
                                                                        <xsl:choose>
                                                                            <xsl:when test="$toSystem = 'DD'">
                                                                                <xsl:value-of select="concat($longD, '.', substring-after((number(concat($longM,'.',$longm))div 60),'.'))" />
                                                                            </xsl:when>
                                                                        </xsl:choose>
                                                                    </xsl:when>
                                                                    <!--ALREADY IN FORMAT REQUIRED-->
                                                                    <xsl:otherwise>
                                                                        <xsl:choose>
                                                                            <xsl:when test="$toSystem = 'DD'">
                                                                                <xsl:value-of select="concat($longD,'.',substring-before($longMm,'.'),substring-after($longMm,'.'))" />
                                                                            </xsl:when>
                                                                        </xsl:choose>
                                                                    </xsl:otherwise>
                                                                </xsl:choose>
								<!--END CONVERT LONGITUDE FOR RMC-->
							    </LONG>
                                                            <TIMESTAMP>
                                                                <xsl:value-of select="concat($hours,':', $mins, ':', $secs)" />
                                                                <!-- IF WE WANTED SECONDS FROM MIDNIGHT - LIKE TAIP THEN
                                                                <xsl:variable name="timeinSec" select="($hours * 3600) + ($mins * 60) + $secs" />
                                                                <xsl:value-of select="$timeinSec"/>
                                                                -->
                                                            </TIMESTAMP>
							</MDBGPS>
						    </rawMessage>
						</xsl:when>
						<xsl:otherwise>
							<!-- Data is invalid, go ahead and discard -->
							<MessageStatus>3</MessageStatus>
							<Type>1</Type>
							<rawMessage>Data status [<xsl:value-of select="$status" />]is invalid!</rawMessage>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:when>
				<xsl:otherwise>
                                    <MessageStatus>3</MessageStatus>
                                    <rawMessage>NOT A KNOWN TRANSACTION TYPE</rawMessage>
                                    <Type>1</Type>
				</xsl:otherwise>
			</xsl:choose>
		</GPSTransform>
	</xsl:template>
</xsl:stylesheet>