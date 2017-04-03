<%
'Ensure a new copy of the page every time
RESPONSE.EXPIRES	= -1
RESPONSE.BUFFER	= TRUE

'***************************************
' Copyright (c) 2001-2008 Webfirm
' http://www.webfirm.com.au
' Ph: 1300 304 779
'
' File:				    helpers.asp
' Author:			    Steven Taddei
' Date Created:   16/01/2015
' Modified By:    Steven Taddei
' Date Modified:  16/01/2015
'
' Purpose:
'	Application based helpers
'***************************************

FUNCTION oConvertToDegree(sDecimal,sType)
  DIM oTemp       : SET oTemp = SERVER.CREATEOBJECT("Scripting.Dictionary")
  DIM sDegrees    : sDegrees  = ""
  DIM sMinutes    : sMinutes  = ""
  DIM sSeconds    : sSeconds  = ""
  DIM sDir        : sDir      = ""
  DIM sToString   : sToString = ""

  IF ( bhaveInfo(sDecimal) AND ISNUMERIC(sDecimal) ) THEN
    sDegrees = FIX(sDecimal)
    sMinutes = (sDecimal - sDegrees) * 60

    SELECT CASE UCASE(sType)
      CASE "LON"
        IF ( sDegrees > 0 ) THEN
          sDir = "E"

        ELSE
          sDir = "W"

        END IF
      CASE ELSE
        IF ( sDegrees > 0 ) THEN
          sDir = "N"

        ELSE
          sDir = "S"

        END IF

    END SELECT

    sToString = " " & sDegrees & "° " & FORMATNUMBER(sMinutes,3) & "' " & sDir
  END IF

  'add to params dictionary
  CALL addToDictionatry(oTemp,REPLACE(sDegrees,"-",""),"degrees")
  CALL addToDictionatry(oTemp,REPLACE(sMinutes,"-",""),"minutes")
  CALL addToDictionatry(oTemp,REPLACE(sSeconds,"-",""),"seconds")
  CALL addToDictionatry(oTemp,REPLACE(sToString,"-",""),"string")
  CALL addToDictionatry(oTemp,sDir,"dir")

  SET oConvertToDegree = oTemp
END FUNCTION

FUNCTION iConvertToDecimal(oCoOrd,sType,sConType)
  DIM sDegrees    : sDegrees  = ""
  DIM sMinutes    : sMinutes  = ""
  DIM sSeconds    : sSeconds  = ""
  DIM sDir        : sDir      = ""
  DIM sToString   : sToString = ""
  DIM iDecimal    : iDecimal  = NULL

  'sConType info'
  'DMS to D.d   : sConType  =  DMS-D
  'DMS to DM.m  : sConType  =  DMS-DM
  'DM.m to D.d  : sConType  =  DM-D

  IF ( NOT bHaveInfo(sConType) ) THEN sConType = "DMS-D"

  IF ( bIsDictionary(oCoOrd) ) THEN
    IF ( oCoOrd.EXISTS("degrees") ) THEN sDegrees = oCoOrd.ITEM("degrees")
    IF ( oCoOrd.EXISTS("minutes") ) THEN sMinutes = oCoOrd.ITEM("minutes")
    IF ( oCoOrd.EXISTS("seconds") ) THEN sSeconds = oCoOrd.ITEM("seconds")
    IF ( oCoOrd.EXISTS("string") ) THEN sToString = oCoOrd.ITEM("string")
    IF ( oCoOrd.EXISTS("dir") ) THEN sDir = oCoOrd.ITEM("dir")

    SELECT CASE UCASE(sConType)
      CASE "DMS-DM"
        'Decimal Degrees = Degrees + .d
        IF ( bHaveInfo(oCoOrd.ITEM("minutes")) AND ISNUMERIC(oCoOrd.ITEM("minutes")) ) THEN
          IF ( bHaveInfo(oCoOrd.ITEM("degrees")) AND ISNUMERIC(oCoOrd.ITEM("degrees")) ) THEN
            IF ( bHaveInfo(oCoOrd.ITEM("seconds")) AND ISNUMERIC(oCoOrd.ITEM("seconds")) ) THEN
              iDecimal = oCoOrd.ITEM("degrees") & " " & FORMATNUMBER(( oCoOrd.ITEM("minutes") + (oCoOrd.ITEM("seconds")/60)),7)
            END IF
          END IF
        END IF

        SELECT CASE UCASE(sDir)
          CASE "W","S"
            iDecimal = "-" & iDecimal

        END SELECT

      CASE "DMS-D"
        'IF ( bHaveInfo(oCoOrd.ITEM("minutes")) AND ISNUMERIC(oCoOrd.ITEM("minutes")) ) THEN
        '  iDecimal = (oCoOrd.ITEM("minutes") / 60)

        '  IF ( bHaveInfo(oCoOrd.ITEM("degrees")) AND ISNUMERIC(oCoOrd.ITEM("degrees")) ) THEN
        '    iDecimal = iDecimal + oCoOrd.ITEM("degrees")
        '  END IF
        'END IF
        IF ( bHaveInfo(oCoOrd.ITEM("minutes")) AND ISNUMERIC(oCoOrd.ITEM("minutes")) ) THEN
          IF ( bHaveInfo(oCoOrd.ITEM("degrees")) AND ISNUMERIC(oCoOrd.ITEM("degrees")) ) THEN
            IF ( bHaveInfo(oCoOrd.ITEM("seconds")) AND ISNUMERIC(oCoOrd.ITEM("seconds")) ) THEN
              iDecimal = FORMATNUMBER(oCoOrd.ITEM("degrees") + (oCoOrd.ITEM("minutes")/60) + (oCoOrd.ITEM("seconds")/3600),7)
            END IF
          END IF
        END IF
        SELECT CASE UCASE(sDir)
          CASE "W","S"
            iDecimal = iDecimal - (iDecimal*2)

        END SELECT

    CASE "DM-D"
      IF ( bHaveInfo(oCoOrd.ITEM("minutes")) AND ISNUMERIC(oCoOrd.ITEM("minutes")) ) THEN
        iDecimal = (oCoOrd.ITEM("minutes") / 60)

        IF ( bHaveInfo(oCoOrd.ITEM("degrees")) AND ISNUMERIC(oCoOrd.ITEM("degrees")) ) THEN
          iDecimal = FORMATNUMBER(iDecimal + oCoOrd.ITEM("degrees"),7)
        END IF
      END IF

      SELECT CASE UCASE(sDir)
        CASE "W","S"
          iDecimal = iDecimal - (iDecimal*2)

      END SELECT
    END SELECT
  END IF

  iConvertToDecimal = iDecimal
END FUNCTION
%>
