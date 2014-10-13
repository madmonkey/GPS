<xsl:stylesheet xmlns:xsl = "http://www.w3.org/1999/XSL/Transform" 
                version = "1.0"
                xmlns:msxsl="urn:schemas-microsoft-com:xslt"
                xmlns:my="http://example.com/2008/my"
                exclude-result-prefixes="msxsl my"                
                >
  <xsl:output method = "xml" standalone ="yes" encoding="UTF-8"/>

  <!--CES_NMEA To TAIP PV.xsl (c) SunGard PS Inc-->
  
  <!--CES_NMEA TO TAIP PV STYLESHEET-->
  
  <msxsl:script implements-prefix="my" language="javascript">
    function ascii_value (c,srchASCIICode)
    {
    var i;
    var bContinue;
    var matchPos;
    var asciiChar;
    i =0;
    bContinue = true;
    matchPos = 0;

    if (c.length>0)
    {
    while(bContinue)
    {
    asciiChar = c.charCodeAt(i);
    if (asciiChar==srchASCIICode)
    {
    i = i + 1;
    matchPos = i;
    bContinue = false;
    break;
    }

    i = i + 1;
    if (i == c.length)
    {
    bContinue = false;
    break;
    }
    }
    }
    return matchPos;
    }

    function GetDataPart(data, packetnum, code)
    {
    var retData;
    var posof127;

    if (packetnum=='100' || packetnum == '108')
    {
    retData = data.substring(16);
    }
    else if (packetnum=='103')
    {
    retData = data.substring(19);
    }
    else if(packetnum=='104')
    {
    posof127 = ascii_value(data, '127');
    retData = data.substring(posof127);
    }

    else
    {
    retData = data;
    }

    return retData;
    }




    function GetTimeToFix(data)
    {
    var retValue;
    var bcontinue;
    var i;
    var j;
    var pads;
    var dataArr;
    var tmp;

    pads ='';
    j=0;
    bcontinue = true;
    if(data.substring(0,1)==',')
    {
    retValue = '000000';
    return retValue;
    }

    dataArr = data.split(",");

    tmp = dataArr[0];


    i = 6 - tmp.length;
    if (i>0)
    {
    while (bcontinue)
    {
    pads = pads + '0';
    j = j + 1
    if (j ==i)
    {
    bcontinue = false;
    break;
    }
    }
    retValue = pads + '' + tmp;
    }
    else
    {
    retValue = tmp.substring(0,6);
    }

    return retValue;
    }


    function GetPaddedString(PadLen,PadStr)
    {
    var i;
    var bcontinue;
    var retPad;

    i = 0;
    bcontinue = true;
    retPad = '';

    if (PadLen>0)
    {
    if (PadStr!='')
    {
    while (bcontinue)
    {
    retPad = retPad + PadStr;
    i = i + 1
    if (i == PadLen)
    {
    bcontinue = false;
    break;
    }
    }
    }

    return retPad;
    }
    else
    {
    return retPad;
    }
    }




    function GetFormattedLatOrLong(data,pos)
    {
    var retValue;
    var bcontinue;
    var i;
    var j;
    var pads;
    var dataArr;
    var tmp;
    var indx;
    var leftpart;
    var rightpart;
    var padstr;

    dataArr = data.split(",");

    tmp = dataArr[pos];

    if (tmp.length=0)
    {
    retValue = '+00000.000000';
    return retValue;
    }

    if (tmp.substring(0,1) !='-')
    {
    if(tmp.substring(0,1)!='+')
    {
    tmp = '+' + tmp;
    }
    }

    indx = tmp.indexOf('.');
    if (indx>=0)
    {
    //formatting latitude or longitude when "." is found
    leftpart = tmp.substring(1,indx);
    if ((5-leftpart.length)>0)
    {
    padstr = GetPaddedString(5-leftpart.length,'0');
    leftpart = tmp.substring(0,1) + padstr + leftpart.substring(0);
    }

    rightpart = tmp.substring(indx+1)
    if ((6 - rightpart.length)>0)
    {
    padstr = GetPaddedString(6 - rightpart.length, '0');
    rightpart = rightpart + padstr;
    }

    retValue = leftpart + '.' + rightpart;
    }
    else
    {
    //formatting latitude or longitude when "." is NOT found
    leftpart = tmp.substring(1);
    if ((5-leftpart.length)>0)
    {
    padstr = GetPaddedString(5-leftpart.length,'0');
    leftpart = tmp.substring(0,1) + padstr + leftpart.substring(0);

    retValue = leftpart + '.000000';
    }
    }

    return retValue;
    }


    function GetLatitude(data)
    {
    var retLat;
    var tmpLat;
    tmpLat = GetFormattedLatOrLong(data,1);

    if (tmpLat.length=0)
    {
    retLat = '+00000.000000';
    }
    else
    {
    retLat = tmpLat;
    }

    return retLat;
    }


    function GetLongitude(data)
    {
    var retLong;
    var tmpLong;
    tmpLong = GetFormattedLatOrLong(data,2);

    if (tmpLong.length=0)
    {
    retLong = '+00000.000000';
    }
    else
    {
    retLong = tmpLong;
    }

    return retLong;
    }


    function GetVelocity(data)
    {
    var tmpVel;
    var dataArr;
    var Padstr;
    var indx;
    var Vel;
    var retVal;

    Padstr = '';
    dataArr = data.split(",");
    retVal = '';

    tmpVel = dataArr[3];

    if (tmpVel.length=0)
    {
    return  '000';
    }
    else
    {
    Vel = tmpVel*0.621;
    tmpVel = "" + Vel;
    indx = tmpVel.indexOf('.');

    if (indx>=0)
    {
    tmpVel = tmpVel.substring(0,indx);
    }

    Padstr = GetPaddedString(3-tmpVel.length,'0');
    }

    retVal = Padstr + '' +  tmpVel;
    return retVal;
    }



    function GetBearing(data)
    {
    var tmpBearing;
    var dataArr;
    var Padstr;
    var indx;
    var Vel;
    var retVal;

    Padstr = '';
    dataArr = data.split(",");
    retVal = '';

    tmpBearing = dataArr[4];

    if (tmpBearing.length=0)
    {
    return  '000';
    }
    else
    {
    indx = tmpBearing.indexOf('.');

    if (indx>=0)
    {
    tmpBearing = tmpBearing.substring(0,indx);
    }

    Padstr = GetPaddedString(3-tmpBearing.length,'0');
    }

    retVal = Padstr + '' +  tmpBearing;
    return retVal;
    }






  </msxsl:script>

  
  


  <!--NMEA To TAIP PV.xsl (c) SunGard HTE Inc-->
	
	<!--NMEA TO TAIP PV STYLESHEET-->

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
		<!--<xsl:variable name="unitID" select="'/@#{ID};'"/>-->
		<xsl:variable name="unitID"> 
			<xsl:if test="'CAD400'='/@#{CAD};'">
             			<xsl:value-of select="'/@#{DEVID};'"/>
			</xsl:if>
			<xsl:if test="not('CAD400'='/@#{CAD};')">
				<xsl:value-of select="'/@#{ID};'"/>
			</xsl:if>
		</xsl:variable>
		<xsl:variable name="aliasID" select="'/@#{ALIAS};'"/>
	<xsl:template match="/">
	<xsl:variable name="transaction" select="substring(.,1)"/>
	
	<GPSTransform>
	<xsl:choose>
		<!--CES Trak-CONTROL Customized NMEA TRANSACTION-->
    <xsl:when test = "substring($transaction,6,3) ='100' or substring($transaction,6,3) = '103' or substring($transaction,6,3) = 104 or substring($transaction,6,3) = '108'" >
      <xsl:variable name="mobilenum" select="substring($transaction,1,5)" />
      <xsl:variable name="packetnum" select="substring($transaction,6,3)" />
      <xsl:variable name="basemodemchannel" select="substring($transaction,9,1)" />
      <xsl:variable name="datapart" select="my:GetDataPart($transaction,$packetnum,'127')" />     

      <xsl:variable name="time" select="my:GetTimeToFix($datapart)" />
      <xsl:variable name="hours" select="substring($time,1,2)" />
      <xsl:variable name="mins" select="substring($time,3,2)" />
      <xsl:variable name="secs" select="substring($time,5,2)" />


      <xsl:variable name="latitude" select="my:GetLatitude($datapart)" />
      <xsl:variable name="latD" select="substring($latitude,3,2)" />
      <xsl:variable name="latM" select="substring($latitude,5,2)" />
      <xsl:variable name="latm" select="substring($latitude,8,6)" />
      <xsl:variable name="FormattedLat" select="substring(concat($latD,substring-after((number(concat($latM,'.',$latm))div 60),'.'),'000000000'),1,7)" />
      
      <xsl:variable name="longitude" select="my:GetLongitude($datapart)" />
      <xsl:variable name="longD" select="substring($longitude,2,3)" />
      <xsl:variable name="longM" select="substring($longitude,5,2)" />
      <xsl:variable name="longm" select="substring($longitude,8,6)" />
      <xsl:variable name="FormattedLong" select="substring(concat($longD,substring-after((number(concat($longM,'.',$longm))div 60),'.'),'0000000000'),1,8)" />

      <xsl:variable name="Velocity" select="my:GetVelocity($datapart)" />
      <xsl:variable name="Bearing" select="my:GetBearing($datapart)" />

      <!--0 - processed, 1 - processed warn, 2 - processed error, 3 - error -->
			<xsl:choose>
				<xsl:when test ="string-length($latD)>0">
					<MessageStatus>0</MessageStatus>
					<rawMessage>
						<xsl:value-of select ="string('RPV')"/>
						<xsl:variable name="timeinSec" select = "($hours * 3600) + ($mins * 60) + $secs"/>
						<xsl:variable name="timePad" select = "substring('00000000',1,5-string-length(string($timeinSec)))"/>
						<xsl:value-of select="concat($timePad,$timeinSec)"/>

            <xsl:value-of select ="substring($latitude,1,1)"/>
            
            
<!--CONVERT LATITUDE FOR CES Trak-CONTROL Customized NMEA TRANSACTION-->

            <xsl:value-of select="$FormattedLat" />
            
<!--END CONVERT LATITUDE FOR CES Trak-CONTROL Customized NMEA TRANSACTION-->

            <xsl:value-of select ="substring($longitude,1,1)"/>
            
<!--CONVERT LONGITUDE FOR CES Trak-CONTROL Customized NMEA TRANSACTION-->
            
            <xsl:value-of select="$FormattedLong" />
            
<!--END CONVERT LONGITUDE FOR CES Trak-CONTROL Customized NMEA TRANSACTION-->		
            
            
						<!--xsl:value-of select="string('FFFGGG22')"/-->
            <xsl:value-of select="$Velocity"/>
            <xsl:value-of select="$Bearing"/>
            <xsl:value-of select="string('22')"/>
            
						<xsl:value-of select ="concat(';ID=',$mobilenum)"/>
						<!--xsl:value-of select ="concat(';UID=',$aliasID)"/-->
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