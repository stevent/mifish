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
FUNCTION iReturnNumber(numeric_value)
  iReturnNumber = 0

  'return result
  IF ( bHaveInfo(numeric_value) AND ISNUMERIC(numeric_value) ) THEN iReturnNumber = INT(numeric_value)
END FUNCTION

'-------------------------------------------------------------------------------
' Purpose:  Make sure a number is returned, allows for a default to be passed in
' Inputs:	  Data that needs to be a number
' Returns:	A numeric value
'-------------------------------------------------------------------------------
FUNCTION iReturnDefaultNumber(numeric_value,default_value)
  iReturnDefaultNumber = 0

  IF ( iReturnNumber(numeric_value) > 0 ) THEN
    iReturnDefaultNumber = INT(numeric_value)
  ELSEIF (  iReturnNumber(default_value) > 0 ) THEN
    iReturnDefaultNumber = INT(default_value)
  END IF
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
' Purpose:  Shows a paging navigation if we need one
' Inputs:	  sURL      - URL for the pagination, must have [@IPAGENO@] to be replaced by
'                       new page number
'           iTopPage  - The top page
'           iPageNo   - The current page we are on
'-------------------------------------------------------------------------------
SUB sShowPaging(sURL,iTopPage,iPageNo)
  DIM iCount, iRangeSelect

  IF ( iTopPage > 1 ) THEN
    RESPONSE.WRITE "<div class=""pagination"">" & vbCrLf
    RESPONSE.WRITE "  <ul class=""clearfix"">" & vbCrLf

    'Display each of the Page numbers for quick access
    FOR iCount = 1 TO iTopPage
      ' Determine whether numbers are within range to be displayed
      IF ( ( iPageNo ) < 5 ) THEN
        iRangeSelect = (CINT(iCount) >= CINT(iPageNo+1) - (CINT(iPageNo+1) - 1)) AND (CINT(iCount) < CINT(iPageNo+1) + (11 - CINT(iPageNo+1)))
      ELSE
        iRangeSelect = ((CINT(iCount) >= CINT(iPageNo+1) - 6) AND (CINT(iCount) < CINT(iPageNo+1) + 6))
      END IF

      IF ( iRangeSelect ) THEN
        'Only like to a page number if it's not the current page
        IF ( iPageNo = iCount ) THEN
          RESPONSE.WRITE "    <li><a href=""" & REPLACE(sURL,"[@IPAGENO@]",iCount) & """ class=""active"">" & iCount & "</a></li>" & vbCrLf
        ELSE
          RESPONSE.WRITE "    <li><a href=""" & REPLACE(sURL,"[@IPAGENO@]",iCount) & """>" & iCount & "</a></li>" & vbCrLf
        END IF
      END IF
    NEXT

    RESPONSE.WRITE "  </ul>" & vbCrLf
    RESPONSE.WRITE "</div>" & vbCrLf
  END IF
END SUB

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
    replaceParamNumber = REPLACE(sSQL,setParamNumber(iNum),sParam)
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
%>
