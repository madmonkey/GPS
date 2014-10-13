<xsl:stylesheet xmlns:xsl = "http://www.w3.org/1999/XSL/Transform" version = "1.0" >
  <xsl:output method = "xml" standalone ="yes" encoding="UTF-8"/>
  <!--NMEA To AVL.xsl (c) SunGard HTE Inc-->

  <xsl:variable name="aliasId" select="'/@#{ALIAS};'"/>

  <xsl:template match="/">
    <xsl:variable name="transaction" select="substring-before(.,',')"/>
    <xsl:choose>
      <!--GPGLL TRANSACTION-->
      <xsl:when test = "$transaction ='GPGLL' or $transaction = '$GPGLL'" >
        <xsl:variable name="latDDM" select="substring-before(substring-after(., concat($transaction, ',')), ',')" />
        <xsl:variable name="latDeg" select="substring($latDDM, 1, 2)" />
        <xsl:variable name="latMin" select="substring($latDDM, 3, 7)" />
        <xsl:variable name="latNS" select="substring-before(substring-after(., concat($latDDM, ',')), ',')" />

        <xsl:variable name="longDDM" select="substring-before(substring-after(., concat($latNS, ',')), ',')" />
        <xsl:variable name="longDeg" select="substring($longDDM, 1, 3)" />
        <xsl:variable name="longMin" select="substring($longDDM, 4, 7)" />
        <xsl:variable name="longEW" select="substring-before(substring-after(., concat($longDDM, ',')), ',')" />

        <xsl:variable name="time" select="substring-before(substring-after(., concat($longEW, ',')), ',')" />
        <xsl:variable name="hours" select="substring($time,1,2)" />
        <xsl:variable name="mins" select="substring($time,3,2)" />
        <xsl:variable name="secs" select="substring($time,5,6)" />

        <xsl:variable name="status" select="substring-before(substring-after(., concat($time, ',')), '-')" />

        <!-- Convert for output AVL format -->
        <xsl:variable name="timeInSec" select = "($hours * 3600) + ($mins * 60) + $secs"/>

        <xsl:variable name="latSign">
          <xsl:choose>
            <xsl:when test="$latNS = 'N'">+</xsl:when>
            <xsl:otherwise>-</xsl:otherwise>
          </xsl:choose>
        </xsl:variable>

        <xsl:variable name="longSign">
          <xsl:choose>
            <xsl:when test="$longEW = 'E'">+</xsl:when>
            <xsl:otherwise>-</xsl:otherwise>
          </xsl:choose>
        </xsl:variable>

        <xsl:variable name="latDD" select="number($latDeg) + number($latMin div 60)" />
        <xsl:variable name="longDD" select="number($longDeg) + number($longMin div 60)" />

        <xsl:variable name="adjustedLat" select="concat($latSign, format-number($latDD, '00.0000000'))" />
        <xsl:variable name="adjustedLong" select="concat($longSign, format-number($longDD, '000.000000'))" />

        <xsl:variable name="age">
          <xsl:choose>
            <xsl:when test="$status = 'A'">
              <xsl:text>2</xsl:text>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>1</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>

        <!-- Output AVL format -->
        <GPSTransform>
          <MessageStatus>0</MessageStatus>
          <rawMessage method="xml">
            <SGAVL>
              <GPS>
                <TDS>
                  <xsl:value-of select="$timeInSec" />
                </TDS>
                <LAT>
                  <xsl:value-of select="$adjustedLat" />
                </LAT>
                <LON>
                  <xsl:value-of select="$adjustedLong" />
                </LON>
                <AGE>
                  <xsl:value-of select="$age" />
                </AGE>
              </GPS>
              <ID>
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
      </xsl:when>
      <!--GPGGA TRANSACTION-->
      <xsl:when test = "$transaction ='GPGGA' or $transaction = '$GPGGA'" >
        <xsl:variable name="time" select="substring-before(substring-after(., concat($transaction, ',')), ',')" />
        <xsl:variable name="hours" select="substring($time,1,2)" />
        <xsl:variable name="mins" select="substring($time,3,2)" />
        <xsl:variable name="secs" select="substring($time,5,6)" />

        <xsl:variable name="latDDM" select="substring-before(substring-after(., concat($time, ',')), ',')" />
        <xsl:variable name="latDeg" select="substring($latDDM, 1, 2)" />
        <xsl:variable name="latMin" select="substring($latDDM, 3, 7)" />
        <xsl:variable name="latNS" select="substring-before(substring-after(., concat($latDDM, ',')), ',')" />

        <xsl:variable name="longDDM" select="substring-before(substring-after(., concat($latNS, ',')), ',')" />
        <xsl:variable name="longDeg" select="substring($longDDM, 1, 3)" />
        <xsl:variable name="longMin" select="substring($longDDM, 4, 7)" />
        <xsl:variable name="longEW" select="substring-before(substring-after(., concat($longDDM, ',')), ',')" />

        <xsl:variable name="fix" select="substring-before(substring-after(., concat($longEW, ',')), ',')" />
        <xsl:variable name="satUsed" select="substring-before(substring-after(., concat($fix, ',')), ',')" />
        <xsl:variable name="hdop" select="substring-before(substring-after(., concat($satUsed, ',')), ',')" />
        <xsl:variable name="altitude" select="substring-before(substring-after(., concat($hdop, ',')), ',')" />

        <!-- Convert for output AVL format -->
        <xsl:variable name="timeInSec" select = "($hours * 3600) + ($mins * 60) + $secs"/>

        <xsl:variable name="latSign">
          <xsl:choose>
            <xsl:when test="$latNS = 'N'">+</xsl:when>
            <xsl:otherwise>-</xsl:otherwise>
          </xsl:choose>
        </xsl:variable>

        <xsl:variable name="longSign">
          <xsl:choose>
            <xsl:when test="$longEW = 'E'">+</xsl:when>
            <xsl:otherwise>-</xsl:otherwise>
          </xsl:choose>
        </xsl:variable>

        <xsl:variable name="latDD" select="number($latDeg) + number($latMin div 60)" />
        <xsl:variable name="longDD" select="number($longDeg) + number($longMin div 60)" />

        <xsl:variable name="adjustedLat" select="concat($latSign, format-number($latDD, '00.0000000'))" />
        <xsl:variable name="adjustedLong" select="concat($longSign, format-number($longDD, '000.000000'))" />

        <!-- Output AVL format -->
        <GPSTransform>
          <MessageStatus>0</MessageStatus>
          <rawMessage method="xml">
            <SGAVL ver="1.0">
              <GPS>
                <TDS>
                  <xsl:value-of select="$timeInSec" />
                </TDS>
                <LAT>
                  <xsl:value-of select="$adjustedLat" />
                </LAT>
                <LON>
                  <xsl:value-of select="$adjustedLong" />
                </LON>
                <ALT>
                  <xsl:value-of select="$altitude" />
                </ALT>
              </GPS>
              <ID>
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
      </xsl:when>
      <!--GPRMC TRANSACTION-->
      <xsl:when test = "$transaction ='GPRMC' or $transaction = '$GPRMC'" >
        <xsl:variable name="time" select="substring-before(substring-after(., concat($transaction, ',')), ',')" />
        <xsl:variable name="hours" select="substring($time,1,2)" />
        <xsl:variable name="mins" select="substring($time,3,2)" />
        <xsl:variable name="secs" select="substring($time,5,6)" />

        <xsl:variable name="status" select="substring-before(substring-after(., concat($time, ',')), ',')" />

        <xsl:variable name="latDDM" select="substring-before(substring-after(., concat($status, ',')), ',')" />
        <xsl:variable name="latDeg" select="substring($latDDM, 1, 2)" />
        <xsl:variable name="latMin" select="substring($latDDM, 3, 7)" />
        <xsl:variable name="latNS" select="substring-before(substring-after(., concat($latDDM, ',')), ',')" />

        <xsl:variable name="longDDM" select="substring-before(substring-after(., concat($latNS, ',')), ',')" />
        <xsl:variable name="longDeg" select="substring($longDDM, 1, 3)" />
        <xsl:variable name="longMin" select="substring($longDDM, 4, 7)" />
        <xsl:variable name="longEW" select="substring-before(substring-after(., concat($longDDM, ',')), ',')" />

        <xsl:variable name="knots" select="substring-before(substring-after(.,concat($longEW,',')),',')" />
        <xsl:variable name="direction" select="substring-before(substring-after(.,concat($knots,',')),',')" />

        <!-- Convert for output AVL format -->
        <xsl:variable name="timeInSec" select = "($hours * 3600) + ($mins * 60) + $secs"/>

        <xsl:variable name="mph" select = "format-number($knots * 1.15077945, '000.0')"/>

        <xsl:variable name="latSign">
          <xsl:choose>
            <xsl:when test="$latNS = 'N'">+</xsl:when>
            <xsl:otherwise>-</xsl:otherwise>
          </xsl:choose>
        </xsl:variable>

        <xsl:variable name="longSign">
          <xsl:choose>
            <xsl:when test="$longEW = 'E'">+</xsl:when>
            <xsl:otherwise>-</xsl:otherwise>
          </xsl:choose>
        </xsl:variable>

        <xsl:variable name="latDD" select="number($latDeg) + number($latMin div 60)" />
        <xsl:variable name="longDD" select="number($longDeg) + number($longMin div 60)" />

        <xsl:variable name="adjustedLat" select="concat($latSign, format-number($latDD, '00.0000000'))" />
        <xsl:variable name="adjustedLong" select="concat($longSign, format-number($longDD, '000.000000'))" />

        <xsl:variable name="age">
          <xsl:choose>
            <xsl:when test="$status = 'A'">
              <xsl:text>2</xsl:text>
            </xsl:when>
            <xsl:otherwise>
              <xsl:text>1</xsl:text>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>

        <!-- Output AVL format -->
        <GPSTransform>
          <MessageStatus>0</MessageStatus>
          <rawMessage method="xml">
            <SGAVL ver="1.0">
              <GPS>
                <TDS>
                  <xsl:value-of select="$timeInSec" />
                </TDS>
                <LAT>
                  <xsl:value-of select="$adjustedLat" />
                </LAT>
                <LON>
                  <xsl:value-of select="$adjustedLong" />
                </LON>
                <MPH>
                  <xsl:value-of select="$mph" />
                </MPH>
                <AGE>
                  <xsl:value-of select="$age" />
                </AGE>
                <DIR>
                  <xsl:value-of select="$direction" />
                </DIR>
              </GPS>
              <ID>
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
      </xsl:when>
    </xsl:choose>
  </xsl:template>
</xsl:stylesheet>