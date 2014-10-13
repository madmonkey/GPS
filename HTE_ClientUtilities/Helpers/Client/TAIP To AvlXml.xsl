<xsl:stylesheet xmlns:xsl = "http://www.w3.org/1999/XSL/Transform" version = "1.0" >
  <xsl:output method = "xml" standalone ="yes" encoding="utf-8" />
  <!--TAIP To AVL.xsl (c) SunGard HTE Inc-->

  <xsl:variable name="aliasId" select="'/@#{ALIAS};'"/>

  <xsl:template match="/">
    <xsl:variable name="transaction" select="substring(., 1, 3)" />

    <xsl:if test="$transaction = 'RPV'" >
      <xsl:variable name="timeofday" select="substring(substring-after(., $transaction), 1, 5)" />

      <xsl:variable name="latSign" select="substring(substring-after(., $transaction), 6, 1)" />
      <xsl:variable name="latDegWhole" select="substring(substring-after(., $transaction),7, 2)"/>
      <xsl:variable name="latDegDecimal" select="substring(substring-after(., $transaction),9, 5)"/>

      <xsl:variable name="longSign" select="substring(substring-after(., $transaction), 14, 1)" />
      <xsl:variable name="longDegWhole" select="substring(substring-after(., $transaction),15,3)" />
      <xsl:variable name="longDegDecimal" select="substring(substring-after(., $transaction),18,5)" />

      <xsl:variable name="speed" select="substring(substring-after(., $transaction), 23, 3)" />

      <xsl:variable name="heading" select="substring(substring-after(., $transaction), 26, 3)" />

      <xsl:variable name="ageOfData" select="substring(substring-after(., $transaction), 30, 1) " />

      <xsl:variable name="source" select ="substring(substring-after(., $transaction), 29, 1)" />

      <xsl:variable name="gpsVehicleId" select="substring-before(substring-after(., 'ID='), ';')" />

      <!-- Produce the AVL message with the data from PV messages-->
      <GPSTransform>
        <MessageStatus>0</MessageStatus>
        <rawMessage method="xml">
          <SGAVL VER="1.0">
            <GPS>
              <xsl:if test="$timeofday != 'AAAAA'">
                <TDS>
                  <xsl:value-of select="$timeofday" />
                </TDS>
              </xsl:if>
              <xsl:if test="$latDegWhole != 'BBB.CCCCC'">
                <LAT>
                  <xsl:value-of select="concat($latSign, $latDegWhole, '.', $latDegDecimal)" />
                </LAT>
              </xsl:if>
              <xsl:if test="$longDegWhole != 'DDDD.EEEEE'">
                <LON>
                  <xsl:value-of select="concat($longSign, $longDegWhole, '.', $longDegDecimal)" />
                </LON>
              </xsl:if>
              <xsl:if test="$speed != 'FFF'">
                <MPH>
                  <xsl:value-of select="$speed" />
                </MPH>
              </xsl:if>
              <xsl:if test="$ageOfData != 'I'">
                <AGE>
                  <xsl:value-of select="$ageOfData" />
                </AGE>
              </xsl:if>
              <xsl:if test="$heading != 'GGG'">
                <DIR>
                  <xsl:value-of select="$heading" />
                </DIR>
              </xsl:if>
              <xsl:if test="$source != 'H'">
                <FIX>
                  <xsl:value-of select="$source" />
                </FIX>
              </xsl:if>
            </GPS>
            <ID>
              <xsl:variable name="unitIdExists" select="substring-before(substring-after(., concat($gpsVehicleId, ';')), '=')" />
              <xsl:if test="$unitIdExists != 'UID'">
                <GDV>
                  <xsl:value-of select="$gpsVehicleId" />
                </GDV>
                <xsl:if test="$aliasId != ''">
                  <UNT>
                    <xsl:value-of select="$aliasId" />
                  </UNT>
                </xsl:if>
              </xsl:if>
              <xsl:if test="$unitIdExists = 'UID'">
                <xsl:variable name="unitId">
                  <xsl:choose>
                    <xsl:when test="contains(substring(substring-after(., concat($unitIdExists, '=')), 1),'*')">
                      <xsl:value-of select="substring-before(substring-after(., concat($unitIdExists, '=')), '*')"/>
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="substring(substring-after(., concat($unitIdExists, '=')), 1)"/>
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:variable>
                <CMP>
                  <xsl:value-of select="$gpsVehicleId" />
                </CMP>
                <xsl:if test="$unitId != ''">
                  <UNT>
                    <xsl:value-of select="$unitId" />
                  </UNT>
                </xsl:if>
              </xsl:if>
            </ID>
          </SGAVL>
        </rawMessage>
        <Type>3</Type>
      </GPSTransform>

    </xsl:if>

    <!-- 
            Finds the variables associated with a LN transaction.
            
            LN message format
            [Time of Day, 9 (AAAAA.BBB)][Latitude, 10][Longitude, 11][Altitude, 9][Horizontal speed, 4][Vertical speed, 5]
            [Heading, 4][# of SVs, 2][SV id, 2][IODE, 2][Reserved, 10][GPS type, 1][Age of Data Indicator, 1]
            -->
    <xsl:if test="$transaction = 'RLN'" >

      <xsl:variable name="timeofday" select="substring(substring-after(., $transaction), 1, 8)" />

      <xsl:variable name="latSign" select="substring(substring-after(., $transaction), 9, 1)" />
      <xsl:variable name="latDegWhole" select="substring(substring-after(., $transaction),10,2)" />
      <xsl:variable name="latDegDecimal" select="substring(substring-after(., $transaction),12,7)" />

      <xsl:variable name="longSign" select="substring(substring-after(., $transaction), 19, 1)" />
      <xsl:variable name="longDegWhole" select="substring(substring-after(., $transaction),20,3)" />
      <xsl:variable name="longDegDecimal" select="substring(substring-after(., $transaction),23,7)" />

      <xsl:variable name="altitudeWhole" select="substring(substring-after(., $transaction), 30, 7)" />
      <xsl:variable name="altitudeDecimal" select="substring(substring-after(., $transaction), 37, 2)" />

      <xsl:variable name="horizontalSpeedWhole" select="substring(substring-after(., $transaction), 39, 3)" />
      <xsl:variable name="horizontalSpeedDecimal" select="substring(substring-after(., $transaction), 42, 1)" />

      <xsl:variable name="verticalSpeedWhole" select="substring(substring-after(., $transaction), 43, 4)" />
      <xsl:variable name="verticalSpeedDecimal" select="substring(substring-after(., $transaction), 47, 1)" />

      <xsl:variable name="headingWhole" select="substring(substring-after(., $transaction), 48, 3)" />
      <xsl:variable name="headingDecimal" select="substring(substring-after(., $transaction), 51, 1)" />

      <xsl:variable name="source" select ="substring(substring-after(., $transaction), 79, 1)" />
      <xsl:variable name="ageOfData" select="substring(substring-after(., $transaction), 80, 1)" />

      <xsl:variable name="gpsVehicleId" select="substring-before(substring-after(substring-after(., $transaction), 'ID='), ';')" />

      <!-- Produce the AVL message -->

      <GPSTransform>
        <MessageStatus>0</MessageStatus>
        <rawMessage method="xml">
          <SGAVL VER="1.0">
            <GPS>
              <TDS>
                <xsl:value-of select="$timeofday" />
              </TDS>
              <LAT>
                <xsl:value-of select="concat($latSign, $latDegWhole, '.', $latDegDecimal)" />
              </LAT>
              <LON>
                <xsl:value-of select="concat($longSign, $longDegWhole, '.', $longDegDecimal)" />
              </LON>
              <HMPH>
                <xsl:value-of select="concat($horizontalSpeedWhole, '.', $horizontalSpeedDecimal)" />
              </HMPH>
              <VMPH>
                <xsl:value-of select="concat($verticalSpeedWhole, '.', $verticalSpeedDecimal)" />
              </VMPH>
              <AGE>
                <xsl:value-of select="$ageOfData" />
              </AGE>
              <DIR>
                <xsl:value-of select="concat($headingWhole, '.', $headingDecimal)" />
              </DIR>
              <ALT>
                <xsl:value-of select="concat($altitudeWhole, '.', $altitudeDecimal)" />
              </ALT>
              <FIX>
                <xsl:value-of select="$source" />
              </FIX>
            </GPS>
            <ID>
              <GDV>
                <xsl:value-of select="$gpsVehicleId" />
              </GDV>
              <xsl:if test="$aliasId != ''">
                <UNT>
                  <xsl:value-of select="$aliasId" />
                </UNT>
              </xsl:if>
            </ID>
          </SGAVL>
        </rawMessage>
        <Type>3</Type>
      </GPSTransform>

    </xsl:if>
  </xsl:template>
</xsl:stylesheet>