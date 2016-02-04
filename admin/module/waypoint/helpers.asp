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
' Version:			  1.0
' Author:			    Steven Taddei
' Modified By:		Steven Taddei
' Date Created:	  20/01/2016
' Last Modified:	20/01/2016
'
' Purpose:
'	helpers for waypoints
'***************************************

'-------------------------------------------------------------------------------
' Purpose:  Save waypoint form data
' Inputs:	  sAction - the action being performed
'           oRecord - the waypoint record we are updating the sounder field data for.
'-------------------------------------------------------------------------------
SUB setFromForm(sAction,oRecord)
  DIM oLongitude      : SET oLongitude  = oConvertToDegree(NULL,"lon")
  DIM oLatitude       : SET oLatitude   = oConvertToDegree(NULL,"lat")

  oLongitude.ITEM("degrees")  = c_FormRequest("LongitudeDegrese")
  oLongitude.ITEM("minutes")  = c_FormRequest("LongitudeMinutes")
  oLongitude.ITEM("dir")      = c_FormRequest("LongitudeDir")
  oLatitude.ITEM("degrees")   = c_FormRequest("LatitudeDegrees")
  oLatitude.ITEM("minutes")   = c_FormRequest("LatitudeMinutes")
  oLatitude.ITEM("dir")       = c_FormRequest("LatitudeDir")

  oRecord.SetValue("Title")       = c_FormRequest("Title")
  oRecord.SetValue("Type")        = c_FormRequest("Type")
  oRecord.SetValue("Longitude")   = iConvertToDecimal(oLongitude,"lon")
  oRecord.SetValue("Latitude")    = iConvertToDecimal(oLatitude,"lat")
  oRecord.SetValue("Notes")       = c_FormRequest("Notes")
  oRecord.SetValue("MemberID")    = Member.FieldValue("ID")
END SUB


'-------------------------------------------------------------------------------
' Purpose:  Saves the sounder field data frpm an xml object
' Inputs:	  oXML - The XML File Object we are importing
' Returns:  the Log of what has happened
'-------------------------------------------------------------------------------
FUNCTION ImportRaymarineDragonflyGPX(oXML)
  DIM oWpt, oWptExt, oRecord
  DIM sName, sInternalID, sLog, sLogLine, sSQL
  DIM iCnt

  'add to log'
  sLog  = sLog & "<p>XML Loaded succesfully.</p>"
  iCnt  = 0

  FOR EACH oWpt IN oXMLDoc.documentElement.selectNodes("wpt")
    'default'
    sLogLine = ""

    IF ( NOT oWpt.selectSingleNode("name") IS NOTHING ) THEN sName = oWpt.selectSingleNode("name").text

    IF ( NOT oWpt.selectSingleNode("extensions") IS NOTHING ) THEN
      SET oWptExt = oWpt.selectSingleNode("extensions")

      IF ( NOT oWptExt.selectSingleNode("raymarine:GUID") IS NOTHING ) THEN
        sInternalID = oWptExt.selectSingleNode("raymarine:GUID").text
      END IF
    END IF

    IF ( bHaveInfo(sInternalID) AND bHaveInfo(sInternalID) ) THEN
      sSQL = "SELECT Waypoint.* "
      sSQL = sSQL & "FROM Waypoint, WaypointSounderData "
      sSQl = sSQL & "WHERE Waypoint.ID=WaypointSounderData.WaypointID "
      sSQL = sSQL & "AND SounderFieldID=1 "
      sSQL = sSQL & "AND ("
      sSQL = sSQL & "       ( "
      sSQL = sSQL & "         UCASE(Title)='" & UCASE(sName) & "' AND ( WaypointSounderData.[Value]='' OR ISNULL(WaypointSounderData.[Value]) ) "
      sSQL = sSQL & "       ) "
      sSQL = sSQL & "       OR "
      sSQL = sSQL & "       ( "
      sSQL = sSQL & "         WaypointSounderData.[Value]='" & UCASE(sInternalID) & "' "
      sSQL = sSQL & "       ) "
      sSQL = sSQL & "     ) "

      SET oRecord   = c_Waypoint.run("FindBySQL",sSQL)

      IF ( oRecord IS NOTHING  ) THEN
        SET oRecord   = c_Waypoint.run("New",NULL)

        'add to log line'
        sLogLine = sLogLine & "Waypoint " & sName & " with internal ID " & sInternalID & " was found. Data will be added"
      ELSE
        'add to log line'
        sLogLine = sLogLine & "Waypoint " & sName & " with internal ID " & sInternalID & " was found. Data will be updated"
      END IF

      oRecord.SetValue("Title")       = sName
      oRecord.SetValue("Type")        = 1
      oRecord.SetValue("Longitude")   = oWpt.GetAttribute("lon")
      oRecord.SetValue("Latitude")    = oWpt.GetAttribute("lat")

      IF ( NOT oWpt.selectSingleNode("cmt") IS NOTHING ) THEN
        oRecord.SetValue("Notes")       = oWpt.selectSingleNode("cmt").text
      END IF

      oRecord.SetValue("MemberID")    = Member.FieldValue("ID")

      IF ( oRecord.run("SaveRecord",NULL) ) THEN
        'add to log line'
        sLogLine = sLogLine & "<br /> - Was succesfully saved. Move on to adding specific ray marine fields."

        CALL saveSounderXML(oRecord,oWpt,oWptExt)
      ELSE
        'add to log line'
        sLogLine = sLogLine & "<br /> - Was not saved. There was an error attempting to add the waypoint data."
      END IF
    ELSE
      'add to log line'
      sLogLine = sLogLine & "Issue with the name (" & sName & ") and internal id (" & sInternalID & ")."
    END IF

    'add to log'
    sLog = sLog & "<p>" & sLogLine & "</p>"

    iCnt = iCnt + 1
  NEXT

  ImportRaymarineDragonflyGPX = sLog
END FUNCTION

'-------------------------------------------------------------------------------
' Purpose:  Saves the sounder field data frpm an xml object
' Inputs:	  oRecord - the waypoint record we are updating the sounder field data for.
'           oWpt    - The initial waypoint xml data
'           oWptExt - The more specific raymarine data
'-------------------------------------------------------------------------------
SUB saveSounderXML(oRecord,oWpt,oWptExt)
  DIM oSounders : SET oSounders = c_Sounder.run("FindByMember",Member.FieldValue("ID"))
  DIM oSounder, oSounderFields, oField
  DIM sFieldValue
  DIM bSaved

  FOR EACH oSounder IN oSounders.ITEMS
    SET oSounderFields = c_SounderField.run("FindBySounder",oSetParams(ARRAY("iSounderID[" & oSounder.FieldValue("ID") & "]","iWaypointID[" & oRecord.FieldValue("ID") & "]")))

    FOR EACH oField IN oSounderFields.ITEMS
      sFieldValue   = ""

      IF ( NOT oWpt.selectSingleNode(oField.FieldValue("FieldName")) IS NOTHING ) THEN
        sFieldValue = oWpt.selectSingleNode(oField.FieldValue("FieldName")).text

      ELSE
        IF ( NOT oWptExt.selectSingleNode(oField.FieldValue("FieldName")) IS NOTHING ) THEN
          sFieldValue = oWptExt.selectSingleNode(oField.FieldValue("FieldName")).text
        END IF

      END IF

      'Check if its linked to a waypoint field
      IF ( bHaveInfo(oField.FieldValue("WaypointField")) ) THEN
        sFieldValue = oRecord.FieldValue(oField.FieldValue("WaypointField"))
      END IF

      oField.SetValue("FieldValue") = sFieldValue

      'only save if its not linked to a waypoint field
      IF ( NOT bHaveInfo(oField.FieldValue("WaypointField")) ) THEN
        bSaved = oField.run("SaveWaypointData",oSetParams(ARRAY("iSounderFieldID[" & oField.FieldValue("ID") & "]","iSounderID[" & oSounder.FieldValue("ID") & "]","iWaypointID[" & oRecord.FieldValue("ID") & "]")))
      END IF
    NEXT
  NEXT
END SUB

'-------------------------------------------------------------------------------
' Purpose:  Saves the sounder fields for a paticulr waypoint. It either adds the
'           field or creates it. It saves from the form.
' Inputs:	  sAction - the action being performed
'           oRecord - the waypoint record we are updating the sounder field data for.
'-------------------------------------------------------------------------------
SUB saveSounderFields(sAction,oRecord)
  DIM oSounders       : SET oSounders   = c_Sounder.run("FindByMember",Member.FieldValue("ID"))
  DIM oSounder, oSounderFields, oField
  DIM sFieldName, sFieldLabel, sFieldValue, sFieldID
  DIM bSaved

  FOR EACH oSounder IN oSounders.ITEMS
    SET oSounderFields = c_SounderField.run("FindBySounder",oSetParams(ARRAY("iSounderID[" & oSounder.FieldValue("ID") & "]","iWaypointID[" & oRecord.FieldValue("ID") & "]")))

    FOR EACH oField IN oSounderFields.ITEMS
      sFieldName    = "SounderField_" & oField.FieldValue("ID")
      sFieldLabel   = oField.FieldValue("Name")
      sFieldValue   = c_FormRequest(sFieldName)
      sFieldID      = "SounderField_" & oField.FieldValue("ID")

      'Check if its linked to a waypoint field
      IF ( bHaveInfo(oField.FieldValue("WaypointField")) ) THEN
        sFieldValue = oRecord.FieldValue(oField.FieldValue("WaypointField"))
      END IF

      oField.SetValue("FieldValue") = sFieldValue

      'only save if its not linked to a waypoint field
      IF ( NOT bHaveInfo(oField.FieldValue("WaypointField")) ) THEN
        bSaved = oField.run("SaveWaypointData",oSetParams(ARRAY("iSounderFieldID[" & oField.FieldValue("ID") & "]","iSounderID[" & oSounder.FieldValue("ID") & "]","iWaypointID[" & oRecord.FieldValue("ID") & "]")))
      END IF
    NEXT
  NEXT
END SUB

'-------------------------------------------------------------------------------
' Purpose:  Creates the gpx file for the raymarine dragonfly
' Inputs:	  oWaypoints - The waypoints we are exporting
'           iSounderID - ID of the sounder we are looking for
'-------------------------------------------------------------------------------
SUB createRaymarineDragonflyExport(oWaypoints,iSounderID)
  DIM sFileName, sExport
  DIM oFSO, oFile, oUpload, oExportFields, oField, oWpt, oSounderFields

  'set file name
  sFileName = "waypoints-" & Member.FieldValue("ID") & ".gpx"

  'Prepare to open text file for reading
  SET oFSO = CREATEOBJECT("SCRIPTING.FILESYSTEMOBJECT")

  'Create the file that will store our valid results
  '											          Location																                          'Overwrite		'Unicode
  SET oFile = oFSO.CREATETEXTFILE(SERVER.MAPPATH(APPLICATION("WaypointsUpload")) & "\" & sFileName, TRUE,			    FALSE)

  oFile.WRITELINE "<?xml version=""1.0"" encoding=""UTF-8""?>"
  oFile.WRITELINE "<gpx xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" version=""1.1"" xmlns=""http://www.topografix.com/GPX/1/1"" creator=""Raymarine"" xmlns:raymarine=""http://www.raymarine.com"" xsi:schemaLocation=""http://www.raymarine.com http://raymarine.com/gpx_schema/RaymarineGPXExtensions.xsd http://www.topografix.com/GPX/1/1 http://www.topografix.com/GPX/1/1/gpx.xsd http://www.topografix.com/GPX/Private/TopoGrafix/0/1 http://www.topografix.com/GPX/Private/TopoGrafix/0/1/topografix.xsd"" xmlns:topografix=""http://www.topografix.com/GPX/Private/TopoGrafix/0/1"">"

  FOR EACH oWpt IN oWaypoints.ITEMS
    SET oSounderFields = c_SounderField.run("FindBySounder",oSetParams(ARRAY("iSounderID[" & iSounderID & "]","iWaypointID[" & oWpt.FieldValue("ID") & "]")))

    'create export fields'
    SET oExportFields = SERVER.CREATEOBJECT("Scripting.Dictionary")

    FOR EACH oField IN oSounderFields.ITEMS
      'only show if its not linked to a waypoint field
      IF ( bHaveInfo(oField.FieldValue("WaypointField")) ) THEN
        CALL addToDictionatry(oExportFields,oWpt.FieldValue(oField.FieldValue("WaypointField")),LCASE(oField.FieldValue("FieldName")))
      ELSE
        CALL addToDictionatry(oExportFields,sReturnDefaultString(oField.FieldValue("FieldValue"),""),LCASE(oField.FieldValue("FieldName")))
      END IF
    NEXT

    oFile.WRITELINE "<wpt lon=""" & oWpt.FieldValue("Longitude") & """ lat=""" & oWpt.FieldValue("Latitude") & """>"
    oFile.WRITELINE " <time>" & oExportFields.ITEM("time") & "</time>"
    oFile.WRITELINE " <name>" & oExportFields.ITEM("name") & "</name>"
    oFile.WRITELINE " <cmt>" & oExportFields.ITEM("cmt") & "</cmt>"
    oFile.WRITELINE " <sym>Fish</sym>"
    oFile.WRITELINE " <extensions>"
    oFile.WRITELINE "   <raymarine:GUID>" & oExportFields.ITEM("raymarine:guid") & "</raymarine:GUID>"
    oFile.WRITELINE "   <raymarine:GroupName>" & oExportFields.ITEM("raymarine:groupname") & "</raymarine:GroupName>"
    oFile.WRITELINE "   <raymarine:GroupGUID>" & oExportFields.ITEM("raymarine:groupguid") & "</raymarine:GroupGUID>"
    oFile.WRITELINE "   <raymarine:WaterTemp>" & oExportFields.ITEM("raymarine:watertemp") & "</raymarine:WaterTemp>"
    oFile.WRITELINE "   <raymarine:WaterDepth>" & oExportFields.ITEM("raymarine:waterdepth") & "</raymarine:WaterDepth>"
    oFile.WRITELINE " </extensions>"
    oFile.WRITELINE "</wpt>"
  NEXT

  oFile.WRITELINE "</gpx>"

  'sExport = oFile.READALL

  'Release the filesystem resources
  oFile.CLOSE
  SET oFile = NOTHING

  'Create our upload object
  SET oUpload = SERVER.CREATEOBJECT("Persits.Upload")

  RESPONSE.CONTENTTYPE = "application/gpx+xml"
  RESPONSE.ADDHEADER "Content-Disposition", "filename=" & sFileName

  'Send the export information to the user
  oUpload.SENDBINARY SERVER.MAPPATH(APPLICATION("WaypointsUpload")) & "\" & sFileName, FALSE

  'Release our asp upload object
  SET oUpload = NOTHING

  'Check to see if the filename that we have selected already exists on the drive
  IF ( oFSO.FILEEXISTS(SERVER.MAPPATH(APPLICATION("WaypointsUpload")) & "\" & sFileName) ) THEN
    oFSO.DELETEFILE(SERVER.MAPPATH(APPLICATION("WaypointsUpload")) & "\" & sFileName)
  END IF

  'Release our filesystem resources
  SET oFSO = NOTHING
END SUB

SUB createJsonDataFile()
  DIM sFileName, sExport, sComma, sLineText
  DIM oFSO, oFile, oUpload, oExportFields, oField, oWpt, oSounderFields
  DIM oParams, oWaypoints, oTypes
  DIM iCount
  DIM sType

  SET oParams = oSetParams(ARRAY("conditions[]","order[]"))

  oParams.ITEM("conditions")  = ARRAY("MemberID = " & setParamNumber(1),Member.FieldValue("ID"))
  oParams.ITEM("order")       = ARRAY("Title ASC")

  SET oWaypoints  = c_Waypoint.run("FindByMember",oParams)
  SET oTypes      = c_WaypointType.run("FindAll",NULL)

  'set file name
  sFileName = MemberWptJsonFilename()

  'Prepare to open text file for reading
  SET oFSO = CREATEOBJECT("SCRIPTING.FILESYSTEMOBJECT")

  'Create the file that will store our valid results
  '						Location																                         'Overwrite		  'Unicode
  SET oFile = oFSO.CREATETEXTFILE(APPLICATION("CacheStore") & "\" & sFileName, TRUE,			    FALSE)

  oFile.WRITELINE "data = {"
  oFile.WRITELINE " ""count"": " & oWaypoints.COUNT & ","
  oFile.WRITELINE " ""items"": ["

  iCount = 0

  FOR EACH oWpt IN oWaypoints.ITEMS
    iCount = iCount + 1

    IF ( iCount = oWaypoints.COUNT ) THEN
      sComma = ""
    ELSE
      sComma = ","
    END IF

    'default
    sType = ""

    IF ( oTypes.EXISTS(CSTR(oWpt.FieldValue("Type"))) ) THEN
      sType = oTypes.ITEM(CSTR(oWpt.FieldValue("Type"))).FieldValue("Name")
    END IF

    'create line'
    sLineText = "{"
    sLineText = sLineText & """id"": """ & oWpt.FieldValue("ID") & ""","
    sLineText = sLineText & """title"": """ & oWpt.FieldValue("Title") & ""","
    sLineText = sLineText & """lon"": """ & oWpt.FieldValue("Longitude") & ""","
    sLineText = sLineText & """lat"": """ & oWpt.FieldValue("Latitude") & ""","
    sLineText = sLineText & """notes"": """ & oWpt.FieldValue("Notes") & ""","
    sLineText = sLineText & """type"": """ & sType & ""","
    sLineText = sLineText & "}" & sComma

    'write line text
    oFile.WRITELINE sLineText
  NEXT

  oFile.WRITELINE " ]"
  oFile.WRITELINE "}"

  'Release the filesystem resources
  oFile.CLOSE
  SET oFile = NOTHING

  'Release our filesystem resources
  SET oFSO = NOTHING
END SUB

FUNCTION MemberWptJsonFilename()
  MemberWptJsonFilename = "waypoints-m" & Member.FieldValue("ID") & ".json"
END FUNCTION
%>
