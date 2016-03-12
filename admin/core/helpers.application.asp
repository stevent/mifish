<%
'Ensure a new copy of the page every time
RESPONSE.EXPIRES	= -1
RESPONSE.BUFFER	= TRUE

'***************************************
' Copyright (c) 2001-2008 Webfirm
' http://www.webfirm.com.au
' Ph: 1300 304 779
'
' File:				    helpers.application.asp
' Author:			    Steven Taddei
' Date Created:   18/06/2012
' Modified By:    Steven Taddei
' Date Modified:  18/06/2012
'
' Purpose:
'	The basic helpers for the application
'***************************************

'-------------------------------------------------------------------------------
' Purpose:  Checks to see if a particular string contains information or not
' Inputs:	  The string to be checked
' Returns:	A boolean value
'-------------------------------------------------------------------------------
FUNCTION bHaveInfo(data)
  'Check to see if our string contains any information
  IF ( NOT ISNULL(TRIM(data)) AND LEN(TRIM(data)) > 0 ) THEN
    bHaveInfo = TRUE
  ELSE
    bHaveInfo = FALSE
  END IF
END FUNCTION

'-------------------------------------------------------------------------------
' Purpose:  Make sure a string is returned
' Inputs:	  Data that needs to be a string
' Returns:	A String
'-------------------------------------------------------------------------------
FUNCTION sReturnString(string_value)
  sReturnString = ""

  'return result
  IF ( bHaveInfo(string_value) ) THEN sReturnString = CSTR(string_value)
END FUNCTION

'-------------------------------------------------------------------------------
' Purpose:  Make sure a string is returned, f it is empty use a default
' Inputs:	  Data that needs to be a string as well as a default
' Returns:	A String
'-------------------------------------------------------------------------------
FUNCTION sReturnDefaultString(string_value, default_string)
  sReturnDefaultString = ""

  'return result
  IF ( bHaveInfo(string_value) ) THEN
    sReturnDefaultString = CSTR(string_value)
  ELSEIF ( bHaveInfo(default_string) ) THEN
    sReturnDefaultString = CSTR(default_string)
  END IF
END FUNCTION

'-------------------------------------------------------------------------------
' Purpose:  Make sure null is returned if the data is empty
' Inputs:	  Data we want to return
' Returns:	Null of a string
'-------------------------------------------------------------------------------
FUNCTION sReturnNull(string_value)
  sReturnNull = NULL

  'return result
  IF ( bHaveInfo(string_value) ) THEN sReturnNull = CSTR(string_value)
END FUNCTION

'-------------------------------------------------------------------------------
' Purpose:  Make sure a boolean value is returned
' Inputs:	  Data that needs to be boolean
' Returns:	A boolean value
'-------------------------------------------------------------------------------
FUNCTION bReturnBoolean(boolean_value)
  bReturnBoolean = FALSE

  'return result
  IF ( bHaveInfo(boolean_value) ) THEN
    SELECT CASE UCASE(boolean_value)
      CASE "TRUE","YES","1"
        bReturnBoolean = TRUE
    END SELECT
  END IF
END FUNCTION

'-------------------------------------------------------------------------------
' Purpose:  Make sure a value is boolean
' Inputs:	  data to check
' Returns:	A boolean value
'-------------------------------------------------------------------------------
FUNCTION bIsBoolnea(boolean_value)
  bIsBoolnea = FALSE

  'return result
  IF ( bHaveInfo(boolean_value) ) THEN
    SELECT CASE UCASE(boolean_value)
      CASE "TRUE","YES","1","FALSE","NO","0"
        bIsBoolnea = TRUE
    END SELECT
  END IF
END FUNCTION

'-------------------------------------------------------------------------------
' Purpose:  Make sure a number is returned
' Inputs:	  Data that needs to be a number
' Returns:	A numeric value
'-------------------------------------------------------------------------------
FUNCTION iReturnInt(numeric_value)
  'return result
  iReturnInt = iReturnDefaultInt(numeric_value,0)
END FUNCTION

'-------------------------------------------------------------------------------
' Purpose:  Make sure a number is returned, allows for a default to be passed in
' Inputs:	  Data that needs to be a number
' Returns:	A numeric value
'-------------------------------------------------------------------------------
FUNCTION iReturnDefaultInt(numeric_value,default_value)
  DIM iTemp : iTemp = 0

  IF ( bHaveInfo(default_value) AND ISNUMERIC(default_value) ) THEN iTemp = default_value

  IF ( bHaveInfo(numeric_value) AND ISNUMERIC(numeric_value) ) THEN
    iTemp = INT(numeric_value)
  END IF

  iReturnDefaultInt = iTemp
END FUNCTION

'-------------------------------------------------------------------------------
' Purpose:  grabs a textual representation of the field type
' Inputs:	  field type id
' Returns:	empty string or a string describing data type
'-------------------------------------------------------------------------------
FUNCTION sGetFieldType(field_type_id)
  sGetFieldType = ""

  IF ( bHaveInfo(field_type_id) ) THEN
    ' Evaluate data type and handle appropriately
    SELECT CASE field_type_id
      CASE adBoolean, adUnsignedTinyInt 'Boolean
        sGetFieldType = "boolean"

      CASE adBinary, adVarBinary, adLongVarBinary 'Binary
        sGetFieldType = "binary"

      CASE adLongVarChar, adLongVarWChar 'Memo
        sGetFieldType = "text"

      CASE adDate, adDBDate, adDBTime, adDBTimeStamp 'Date/Time
        sGetFieldType = "date_time"

      CASE adVarChar, adVarWChar 'var char
        sGetFieldType = "varchar"

      CASE adSmallInt 'smallint
        sGetFieldType = "smallint"

      CASE adInteger 'int
        sGetFieldType = "int"

      CASE adDouble 'float
        sGetFieldType = "float"

      CASE adCurrency 'currency
        sGetFieldType = "currency"

    END SELECT
  END IF
END FUNCTION

'-------------------------------------------------------------------------------
' Based on a script found at 4 guys from rolla
'     http://www.4guysfromrolla.com/ASPScripts/PrintPage.asp?REF=%2Fwebtech%2F022504-1.shtml
' Purpose:  returns a file as a useable string
' Inputs:	  web file path we want to return
' Returns:	NULL or string
'-------------------------------------------------------------------------------
FUNCTION getMappedFileAsString(byVal strFilename)
  DIM oFSO : SET oFSO = SERVER.CREATEOBJECT("Scripting.FilesystemObject")
  DIM oTextStream

  getMappedFileAsString = NULL

  IF ( oFSO.FILEEXISTS(SERVER.MAPPATH(strFilename)) ) THEN
    SET oTextStream = oFSO.OPENTEXTFILE(SERVER.MAPPATH(strFilename), 1)

    getMappedFileAsString = oTextStream.READALL

    oTextStream.CLOSE

    SET oTextStream = NOTHING
  END IF

  SET oFSO = NOTHING
END FUNCTION

'-------------------------------------------------------------------------------
' Based on a script found at 4 guys from rolla
'     http://www.4guysfromrolla.com/ASPScripts/PrintPage.asp?REF=%2Fwebtech%2F022504-1.shtml
' Purpose:  fixes a string so it can be exected in ASP code
' Inputs:	  string content
' Returns:	string
'-------------------------------------------------------------------------------
FUNCTION fixInclude(content)
  DIM sOutput : sOutput = ""
  DIM iPost1  : iPost1  = INSTR(content,"<%")
  DIM iPost2  : iPost2  = INSTR(content,"%"& ">")
  DIM sBefore, sMiddle, sAfter

  'if there exists aspStartTag
  IF ( iPost1 > 0 ) THEN
    'text content  before aspStartTag
    sBefore = MID(content,1,iPost1-1)

    'remove linebreaks
    sBefore = replace(sBefore,vbcrlf,"")

    'put content into a response.write
    sBefore = vbcrlf & "RESPONSE.WRITE "" " & sBefore & """" &vbcrlf

    'get code content between aspStartTag  and  aspEndTag
    sMiddle = MID(content,iPost1+2,(iPost2-iPost1-2))

    'get text content after aspEndTag
    sAfter = MID(content,iPost2+2,LEN(content))

    'recurse through the rest
    sOutput = sBefore & sMiddle & fixInclude(sAfter)

  'did not find any aspStartTags
  ELSE
    'remove linebreaks in file
    content = replace(content,vbcrlf,"")
    sOutput = vbcrlf & "RESPONSE.WRITE""" & content &""""
  END IF

  fixInclude = sOutput
END FUNCTION

'-------------------------------------------------------------------------------
' Purpose:  grabs an include file to be executed in the code. "Hacky" wa around
' dynamic includes for classic ASP.
' Inputs:	  web file path we want to return
' Returns:	NULL or string
'-------------------------------------------------------------------------------
FUNCTION grabInclude(strFilename)
  DIM sContent : sContent = getMappedFileAsString(strFilename)

  IF ( bHaveInfo(sContent) ) THEN
    sContent = fixInclude(sContent)
  END IF

  grabInclude = sContent
END FUNCTION

'-------------------------------------------------------------------------------
' Based on a script found at 4 guys from rolla
'     http://www.4guysfromrolla.com/ASPScripts/PrintPage.asp?REF=%2Fwebtech%2F022504-1.shtml
' Purpose:  returns a file as a useable string
' Inputs:	  web file path we want to return
' Returns:	NULL or string
'-------------------------------------------------------------------------------
FUNCTION bFileExists(strFilename)
  DIM oFSO : SET oFSO = SERVER.CREATEOBJECT("Scripting.FilesystemObject")

  IF ( bHaveInfo(strFilename) ) THEN
    bFileExists = oFSO.FILEEXISTS(SERVER.MAPPATH(strFilename))
  ELSE
    bFileExists = FALSE
  END IF

  SET oFSO = NOTHING
END FUNCTION

FUNCTION sGenerateSQL(sTable,aArray)
  DIM sSQL      : sSQL      = "SELECT * FROM " & sTable
  DIM sOptions  : sOptions  = ""
  DIM sSortBy   : sSortBy   = ""
  DIM aOptions
  DIM iCount, iOptionCount
  DIM sField, sOperator, sValue

  IF ( ISARRAY(aArray) ) THEN
    FOR iCount = 0 TO UBOUND(aArray)
      aOptions = aArray(iCount)

      IF ( ISARRAY(aOptions) ) THEN
        sField    = ""
        sOperator = ""
        sValue    = ""

        FOR iOptionCount = 0 TO UBOUND(aOptions)
          SELECT CASE INT(iOptionCount)
            CASE 0
              sField = aOptions(iOptionCount)

            CASE 1
              sOperator = aOptions(iOptionCount)

            CASE 2
              sValue = aOptions(iOptionCount)

          END SELECT
        NEXT

        IF ( bHaveInfo(sField) AND bhaveInfo(sValue) ) THEN
          IF ( UCASE(sField) = "ORDER BY" ) THEN
            IF ( bHaveInfo(sSortBy) ) THEN sSortBy = sSortBy & ", "
            sSortBy = sSortBy & sValue
            IF ( bHaveInfo(sOperator) ) THEN sSortBy = sSortBy & " " & sOperator

          ELSEIF ( bHaveInfo(sOperator) ) THEN
            IF ( bHaveInfo(sOptions) ) THEN sOptions = sOptions & " AND "
            sOptions = sOptions & sField & " " & sOperator & " " & sValue

          END IF
        END IF
      END IF
    NEXT
  END IF

  'add any options and sort bys
  IF ( bHaveInfo(sOptions) ) THEN sSQL = sSQL & " WHERE " & sOptions
  IF ( bHaveInfo(sSortBy) ) THEN sSQL = sSQL & " ORDER BY " & sSortBy

  sGenerateSQL = sSQL
END FUNCTION

FUNCTION NameFromField(sField)
  DIM sName

  sName = REPLACE(sField,"_"," ")

  NameFromField = sName
END FUNCTION

FUNCTION bFoundInArray(sString,aArray)
  DIM bTemp   : bTemp   = FALSE
  DIM iCount  : iCount  = 0

  IF ( ISARRAY(aArray) AND bhaveInfo(sString) ) THEN
    DO WHILE ( iCount <= UBOUND(aArray) AND NOT bTemp )
      IF ( UCASE(aArray(iCount)) = UCASE(sString) ) THEN bTemp = TRUE

      iCount = iCount + 1
    LOOP
  END IF

  bFoundInArray = bTemp
END FUNCTION


'-------------------------------------------------------------------------------
' Purpose:  checks to see if a function exists
' Inputs:	  sFunction - The function we are checking
' Returns:  TRUE|FALSE
'-------------------------------------------------------------------------------
FUNCTION bFunctionExists(sFunction)
  DIM test

  ON ERROR RESUME NEXT

  EXECUTE("test = " & sFunction)

  bFunctionExists = (ERR.NUMBER = 0)
END FUNCTION


'-------------------------------------------------------------------------------
' Purpose:  Retrieve the URL of the current page, including rewritten URL
'				addresses and querystring variables if applicable
' Inputs:	None
' Returns:	The URL of the current page
'-------------------------------------------------------------------------------
FUNCTION sGetCurrentURL()
  DIM sURL

  'Check to see if we have a rewritten URL
  IF ( bHaveInfo(REQUEST.SERVERVARIABLES("HTTP_X_REWRITE_URL")) AND REQUEST.SERVERVARIABLES("HTTP_X_REWRITE_URL") <> "/" ) THEN
    'Retrieve our re-written URL
    sURL = REQUEST.SERVERVARIABLES("HTTP_X_REWRITE_URL")
  ELSE
    'Retrieve our page URL
    sURL = REQUEST.SERVERVARIABLES("URL")

    'Check to see if we have any query string variables
    IF ( bHaveInfo(REQUEST.SERVERVARIABLES("QUERY_STRING")) ) THEN
      sURL = sURL & "?" & REQUEST.SERVERVARIABLES("QUERY_STRING")
    END IF
  END IF

  sGetCurrentURL = sURL
END FUNCTION

'-------------------------------------------------------------------------------
' Purpose:  Capitalises first letter of every string, makes the rest lowercase
' Inputs:	  sData: The data we are capetalizing
' Returns:	any returned data
'-------------------------------------------------------------------------------
FUNCTION sCapitalizeWord(sData)
	DIM sUcase, sLcase, sNewString

	'default
	sNewString = ""

	IF ( bHaveInfo(sData) ) THEN
		sUcase		= UCASE(MID(sData,1,1))
		sLcase		= LCASE(MID(sData,2,LEN(sData)-1))

		sNewString = sUcase & sLcase
	END IF

	'return value
	sCapitalizeWord = sNewString
END FUNCTION

'-------------------------------------------------------------------------------
' Purpose:  Sets a params dictionary from an array
' Inputs:	  aParams: Array we are setting from
' Returns:	dictionaty of params
'-------------------------------------------------------------------------------
FUNCTION oSetParams(aParams)
  DIM iCount
  DIM oKey, oRegEx, oMatch, oMatches
  DIM oParams, oField
  DIM sArray, sParamName, sParamValue
  DIM aTemp

  SET oParams = SERVER.CREATEOBJECT("Scripting.Dictionary")

  IF ( ISARRAY(aParams) ) THEN
    FOR iCount = 0 TO UBOUND(aParams)
      sArray = aParams(iCount)

      IF ( bHaveInfo(sArray) ) THEN
        'craete field object
        SET oField = NEW C_Like_Field

        'grab param name
        aTemp       = SPLIT(sArray,"[")
        sParamName  = aTemp(0)

        'lets find the value
        SET oRegEx = NEW RegExp

        oRegEx.Pattern  = "\[([\w\s]*)\]"
        oRegEx.Global   = TRUE

        SET oMatches = oRegEx.Execute(sArray)

        FOR EACH oMatch IN oMatches
          sParamValue = REPLACE(REPLACE(oMatch,"[",""),"]","")
        NEXT

        'set object details
        oField.Name   = sParamName
        oField.Value  = sParamValue

        'add to params dictionary
        CALL addToDictionatry(oParams,oField,oField.Name)
      END IF
    NEXT
  END IF

  SET oSetParams = oParams
END FUNCTION

'-------------------------------------------------------------------------------
' Purpose:  adds an item to a dictionaty object
' Inputs:	  oDisctionary: Dictionary we are adding to
'           sValue:       value we are adding
'           sKey:         key for the new value
'-------------------------------------------------------------------------------
SUB addToDictionatry(oDisctionary,sValue,sKey)
  IF ( bIsDictionary(oDisctionary) ) THEN
    IF ( oDisctionary.EXISTS(CSTR(sKey)) ) THEN
      oDisctionary.REMOVE(CSTR(sKey))
    END IF

    oDisctionary.ADD CSTR(sKey), sValue
  END IF
END SUB

'-------------------------------------------------------------------------------
' Purpose:  figures out if object passed through is a dictionary object
' Inputs:	  oDict: Dictionary we are checking
' Returns:  TRUE|FALSE
'-------------------------------------------------------------------------------
FUNCTION bIsDictionary(oDict)
  bIsDictionary = ( UCASE(typename(oDict)) = "DICTIONARY" )
END FUNCTION

'-------------------------------------------------------------------------------
' Purpose:  figures out if an object passed through is empty or not
' Inputs:	  oObj: The object we are checking
' Returns:  TRUE|FALSE
'-------------------------------------------------------------------------------
FUNCTION bIsEmpty(oObj)
  bIsEmpty = ( oObj IS NOTHING )
END FUNCTION

'-------------------------------------------------------------------------------
' Purpose:  Adds fields and field properties to a field object
' Inputs:
'-------------------------------------------------------------------------------
SUB addField(rsRecordset,oField,iCount)
  'make sure we have a record before trying to add the value
  IF ( NOT rsRecordSet.RS.EOF ) THEN
    oField.Value = rsRecordSet.RS.FIELDS(iCount).VALUE
  END IF

  'set the properties
  oField.Name          = rsRecordSet.RS.FIELDS(iCount).NAME
  oField.FieldType     = sGetFieldType(rsRecordSet.RS.FIELDS(iCount).TYPE)
  oField.PrimaryKey    = rsRecordSet.RS.FIELDS(iCount).PROPERTIES.ITEM("KEYCOLUMN")
  oField.Unique        = rsRecordSet.RS.FIELDS(iCount).PROPERTIES.ITEM("KEYCOLUMN")

  'set max length property
  SELECT CASE UCASE(oField.FieldType)
    CASE "DATE_TIME"
      oField.MaxSize      = 28

    CASE ELSE
      oField.MaxSize      = rsRecordSet.RS.FIELDS(iCount).DEFINEDSIZE

  END SELECT
END SUB

'-------------------------------------------------------------------------------
' Based on a script found at 4 guys from rolla
'     http://www.4guysfromrolla.com/ASPScripts/PrintPage.asp?REF=%2Fwebtech%2F022504-1.shtml
' Purpose:  returns a file as a useable string
' Inputs:	  web file path we want to return
' Returns:	NULL or string
'-------------------------------------------------------------------------------
FUNCTION file_exists(strFilename)
  DIM oFSO : SET oFSO = SERVER.CREATEOBJECT("Scripting.FilesystemObject")

  IF ( bHaveInfo(strFilename) ) THEN
    file_exists = oFSO.FILEEXISTS(SERVER.MAPPATH(strFilename))
  ELSE
    file_exists = FALSE
  END IF

  SET oFSO = NOTHING
END FUNCTION

'-------------------------------------------------------------------------------
' Purpose:  adds an item to a dictionaty object
' Inputs:	  oDisctionary: Dictionary we are adding to
'           sValue:       value we are adding
'           sKey:         key for the new value
'-------------------------------------------------------------------------------
SUB addToDictionatry(oDisctionary,sValue,sKey)
  IF ( bIsDictionary(oDisctionary) ) THEN
    IF ( oDisctionary.EXISTS(CSTR(sKey)) ) THEN
      oDisctionary.REMOVE(CSTR(sKey))
    END IF

    oDisctionary.ADD CSTR(sKey), sValue
  END IF
END SUB


'-------------------------------------------------------------------------------
' Purpose:  Strips out any characters which are not allowed to be used in the
'				URL of a webpage. It performs a number of opertions on the text
'				string to make it compatible.
' Inputs:	The string to be processed
' Returns:	The processed value
'-------------------------------------------------------------------------------
FUNCTION sFileFriendly(sTemp)
  DIM sRewritten
  DIM oRegExp

  'Create a copy of the incoming variable
  sRewritten = sTemp

  'Check the length before trying to alter the string
  IF ( bHaveInfo(sRewritten) ) THEN
    'Create our regular expression object
    SET oRegExp=NEW RegExp

    WITH oRegExp
      .PATTERN = "[^A-Za-z0-9\-\. ]"
      .IGNORECASE = FALSE
      .GLOBAL = TRUE
    END WITH

    'Strip out all invalid characters
    sRewritten = oRegExp.REPLACE(sRewritten, "")

    'Release our regular expression object resources
    SET oRegExp = NOTHING

    'Replace any spaces that exist with single dashes
    sRewritten = REPLACE(sRewritten," ","-")

    'Remove all double dashes in our filename
    DO WHILE ( INSTR(sRewritten,"--")>0 )
      sRewritten = REPLACE(sRewritten,"--","-")
    LOOP

    'Convert our filename to lowercase
    sRewritten = LCASE(sRewritten)
  END IF

  'All the characters were invalid so we give the file a generic name of "file"
  IF ( INSTR(sRewritten,".") = 1 ) THEN
    sRewritten = "file" & sRewritten
  END IF

  sFileFriendly = sRewritten
END FUNCTION

FUNCTION getCurrentPage()
  DIM sPath : sPath = REQUEST.SERVERVARIABLES("PATH_INFO")
  DIM aPath : aPath = SPLIT(sPath,"/")

  getCurrentPage = aPath(UBOUND(aPath))
END FUNCTION

FUNCTION setParamNumber(iNum)
  setParamNumber = "/c" & iNum & "c/"
END FUNCTION

FUNCTION replaceParamNumber(iNum,sSQL,sParam)
  IF ( bHaveInfo(sSQL) ) THEN
    replaceParamNumber = REPLACE(sSQL,setParamNumber(iNum),sReturnString(sParam))
  END IF
END FUNCTION

FUNCTION setConditions(aConditions)
  DIM sQuery : sQuery = ""
  DIM iLen, iConditionValues, iCount

  IF ( ISARRAY(aConditions) ) THEN
    iLen =  UBOUND(aConditions)

    sQuery = aConditions(0)

    IF ( iLen > 0 ) THEN
      FOR iCount = 1 TO iLen
        sQuery = replaceParamNumber(iCount,sQuery,aConditions(iCount))
      NEXT
    END IF
  END IF

  setConditions = sQuery
END FUNCTION

'-------------------------------------------------------------------------------
' Purpose:  Executes an sql script
' Inputs:   sSQL      - DQL we are using
'-------------------------------------------------------------------------------
SUB executeSQL(sSQL)
  DIM oConn
  DIM rsExecute

  IF ( bhaveInfo(sSQL) ) THEN
    'Opens connection to the database
    SET oConn = createConnection

    'execute connection'
    SET rsExecute = oConn.conn.EXECUTE(sSQL,,&H0001 + &H00000080)
    SET rsExecute = NOTHING

    'Release the database connection
    SET oConn	= NOTHING
  END IF
END SUB

'-------------------------------------------------------------------------------
' Purpose:  To display paging navigation in a pre-determined format.
' Inputs:	  sURL - the url. This will be used as a telplate for
'						 the paging links. Iw will have an identifier [@CP@] so
'						 we can replace it with the page no.
'           iPageSize   - The current page size
'           iCurrentPage     - The current page number we are on
'           iTotalCount - The total amount of records in the recordset
'-------------------------------------------------------------------------------
SUB DisplayPagingNav(sURL,iPageSize,iCurrentPage,iTotalCount)
  DIM iCount, iRangeSelect
  DIM sSpacer

  'do we need to page
  IF ( iTotalCount > iPageSize ) THEN
    'make sure page size has a number
    IF ( NOT ISNUMERIC(iPageSize) OR NOT bHaveInfo(iPageSize) ) THEN
      iPageSize = 6
    ELSE
      iPageSize = INT(iPageSize)
    END IF

    RESPONSE.WRITE "<div class=""pagination"">"
    RESPONSE.WRITE "  <ul>"

    'do we show the previouse
    IF ( NOT iCurrentPage < 2 ) THEN
      RESPONSE.WRITE "  <a class=""button special small prev"" href=""" & REPLACE(sURL,"[@CP@]",(iCurrentPage-1)) & """>Prev</a>"
    END IF

    'Display each of the Page numbers for quick access
    FOR iCount = 1 TO FIX((iTotalCount-1) / iPageSize)+1
      ' Determine whether numbers are within range to be displayed
      IF ( (iCurrentPage ) < 5 ) THEN
        iRangeSelect = (CINT(iCount) >= CINT(iCurrentPage+1) - (CINT(iCurrentPage+1) - 1)) AND (CINT(iCount) < CINT(iCurrentPage+1) + (11 - CINT(iCurrentPage+1)))
      ELSE
        iRangeSelect = ((CINT(iCount) >= CINT(iCurrentPage+1) - 6) AND (CINT(iCount) < CINT(iCurrentPage+1) + 6))
      END IF

      IF ( iRangeSelect ) THEN
        'Only like to a page number if it's not the current page
        IF ( iCurrentPage = iCount ) THEN
          RESPONSE.WRITE "  <a class=""button special small num active"" href=""" & REPLACE(sURL,"[@CP@]",iCount) & """>" & iCount & "</a>"
        ELSE
          RESPONSE.WRITE "  <a class=""button special small num"" href=""" & REPLACE(sURL,"[@CP@]",iCount) & """>" & iCount & "</a>"
        END IF

        'set spacer
        IF ( NOT bHaveInfo(sSpacer) ) THEN sSpacer = "&nbsp;|&nbsp;"
      END IF
    NEXT

    IF ( iCurrentPage < FIX((iTotalCount-1) / iPageSize)+1 ) THEN
      RESPONSE.WRITE "  <a class=""button special small next"" href=""" & REPLACE(sURL,"[@CP@]",(iCurrentPage+1)) & """>Next</a>"
    END IF

    RESPONSE.WRITE "</div>"
  END IF
END SUB

FUNCTION sAddToQueryString(sQueryString,sName,sVal)
  DIM sTemp : sTemp = ""

  IF ( bHaveInfo(sName) AND bhaveInfo(sVal) ) THEN
    IF ( bHaveInfo(sQueryString) ) THEN
      sTemp = sTemp & "&"
    ELSE
      sTemp = sTemp & "?"
    END IF

    sTemp = sTemp & sName & "=" & sVal
  END IF

  sAddToQueryString = sQueryString & sTemp
END FUNCTION

'-------------------------------------------------------------------------------
' Purpose:  To format a date into a specific format.
' Inputs:	sFormat	- the string the indicates the format that is required. See
'				the PHP date() function or the comments below for more information.
'				dDate		- the date that is to be formatted
' Returns:	The formatted date.
'-------------------------------------------------------------------------------
FUNCTION dGetFormattedDate( sFormat, dDate )
  DIM iCount
  DIM bReserved
  DIM sDate, sTemp, sChar
  DIM arrFormat

  'Ensure we have valid paramters before continuing
  IF ( bHaveInfo(sFormat) AND bHaveInfo(dDate) AND ISDATE(dDate) ) THEN
    'Reset the size of our array based on the length of our incoming format string
    REDIM arrFormat(LEN(sFormat))

    'Loop through our formatting string
    FOR iCount=0 TO LEN(sFormat)-1
      'Store this character in our array for later use
      arrFormat(iCount) = MID(sFormat,iCount+1,1)
    NEXT

    'Loop through our array of characters
    FOR iCount=0 TO LEN(sFormat)
      'Reset our variable default
      bReserved = FALSE

      'Retrieve the character from the array
      sChar = arrFormat(iCount)

      'Check to see if we are at the start of the array or not
      IF ( iCount > 0 ) THEN
        'Check to see if this character is a reserved word
        IF ( sChar = "\" ) THEN
          'Set that this character is reserved so that it is not processed
          bReserved = TRUE
        END IF
      END IF

      'Do not process the character as it is reserved
      IF ( NOT bReserved ) THEN
        'Inspect the character to see if it matches what we are looking for
        SELECT CASE sChar
          'Day formatting
          CASE "d": 'Day of the month, 2 digits with leading zeros (01 to 31)
            sChar = sPrefixZero(DAY(dDate))

          CASE "D": 'A textual representation of a day, three letters (Mon through Sun)
            sChar = LEFT(WEEKDAYNAME(WEEKDAY(dDate)),3)

          CASE "j": 'Day of the month without leading zeros (1 to 31)
            sChar = DAY(dDate)

          CASE "l": 'A full textual representation of the day of the week (Sunday through Saturday)
            sChar = WEEKDAYNAME(WEEKDAY(dDate))

          CASE "N": 'ISO-8601 numeric representation of the day of the week (1 for Monday, through 7 for Sunday)
            sChar = WEEKDAY(dDate,2)

          CASE "S": 'English ordinal suffix for the day of the month, 2 characters (st, nd, rd or th. Works well with j)
            'Determine what our correct suffix is based on the final number of our day
            SELECT CASE sPrefixZero(DAY(dDate))
              CASE "01", "21", "31":
                sChar = "st"
              CASE "02", "22":
                sChar = "nd"
              CASE "03", "23":
                sChar = "rd"
              CASE ELSE
                sChar = "th"
            END SELECT

          CASE "w": 'Numeric representation of the day of the week (0 for Sunday, through 6 for Saturday)
            sChar = (WEEKDAY(dDate,1))-1

          CASE "z": 'The day of the year (starting from 0 through 365)
            sChar = DATEDIFF("d",DATESERIAL(YEAR(dDate),"01","01"),dDate)

          'Month formatting
          CASE "F": 'A full textual representation of a month, such as January or March (January through December)
            sChar = MONTHNAME(MONTH(dDate),FALSE)

          CASE "m": 'Numeric representation of a month, with leading zeros (01 through 12)
            sChar = sPrefixZero(MONTH(dDate))

          CASE "M": 'A short textual representation of a month, three letters (Jan through Dec)
            sChar = MONTHNAME(MONTH(dDate),TRUE)

          CASE "n": 'Numeric representation of a month, without leading zeros (1 through 12)
            sChar = MONTH(dDate)

          CASE "t": 'Number of days in the given month (28 through 31)
            sChar = DAY(DATEADD("d",-1,DATEADD("m",1,DATESERIAL(YEAR(dDate),MONTH(dDate),1))))

          'Year formatting
          CASE "L": 'Whether it's a leap year (1 if it is a leap year, 0 otherwise.)s
            'Determine if this year is a leap year or not
            IF ( YEAR(dDate) MOD 400 = 0 ) THEN
              sChar = 1
            ELSEIF ( YEAR(dDate) MOD 100 = 0 ) THEN
              sChar = 0
            ELSEIF ( YEAR(dDate) MOD 4 = 0 ) THEN
              sChar = 1
            ELSE
              sChar = 0
            END IF

          CASE "Y": 'A full numeric representation of a year, 4 digits (Examples: 1999 or 2003)
            sChar = YEAR(dDate)

          CASE "y": 'A two digit representation of a year (Examples: 99 or 03)
            sChar = RIGHT(YEAR(dDate),2)

          'Time formatting
          CASE "a": 'Lowercase Ante meridiem and Post meridiem (am or pm)
            sChar = LCASE(RIGHT(FORMATDATETIME(dDate,3),2))

          CASE "A": 'Uppercase Ante meridiem and Post meridiem (AM or PM)
            sChar = UCASE(RIGHT(FORMATDATETIME(dDate,3),2))

          CASE "g": '12-hour format of an hour without leading zeros (1 through 12)
            'Check to see what hour we currently have
            IF ( HOUR(dDate) > 12 ) THEN
              sChar = HOUR(dDate) - 12
            ELSE
              sChar = HOUR(dDate)
            END IF

          CASE "G": '24-hour format of an hour without leading zeros (0 through 23)
            sChar = HOUR(dDate)

          CASE "h": '12-hour format of an hour with leading zeros (01 through 12)
            'Check to see what hour we currently have
            IF ( HOUR(dDate) > 12 ) THEN
              sChar = sPrefixZero(HOUR(dDate) - 12)
            ELSE
              sChar = sPrefixZero(HOUR(dDate))
            END IF

          CASE "H": '24-hour format of an hour with leading zeros (00 through 23)
            sChar = sPrefixZero(HOUR(dDate))

          CASE "i": 'Minutes with leading zeros (00 to 59)
            sChar = sPrefixZero(MINUTE(dDate))

          CASE "s": 'Seconds, with leading zeros (00 through 59)
            sChar = sPrefixZero(SECOND(dDate))

        END SELECT
      END IF

      'Append this character on to our string
      sDate = sDate & sChar
    NEXT
  ELSE
    sDate = ""
  END IF

  dGetFormattedDate = sDate
END FUNCTION

'-------------------------------------------------------------------------------
' Purpose:  Prefixes a single number with a zero.
' Inputs:	The string/number to be checked.
' Returns:	The same value with a zero appended to the front if required.
'-------------------------------------------------------------------------------
FUNCTION sPrefixZero(sString)
  DIM sTemp

  'Take a copy of the passed in string before continuing
  sTemp = sString

  'Check the length of the string
  IF ( LEN(TRIM(sTemp)) = 1 ) THEN
    sTemp = "0" & sTemp
  END IF

  sPrefixZero = sTemp
END FUNCTION
%>
