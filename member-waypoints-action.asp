<%
OPTION EXPLICIT

'Ensure a new copy of the page every time
RESPONSE.EXPIRES	= -1
RESPONSE.BUFFER	= TRUE

'***************************************
' Copyright (c) 2001-2008 Webfirm
' http://www.webfirm.com.au
' Ph: 1300 304 779
'
' File:				    member-waypoints-action.asp
' Version:			  1.0
' Author:			    Steven Taddei
' Modified By:		Steven Taddei
' Date Created:	  16/01/2016
' Last Modified:	16/01/2016
'
' Purpose:
'	Runs actions for the member waypoint
' section'
'***************************************
%>
<!--#include virtual="/admin/core/core.includes.asp"-->
<%
CALL Member.run("Login",NULL)
CALL Member.run("SecurePage",NULL)

DIM sAction     : sAction       = sReturnDefaultString(c_FormRequest("Action"),"unknown")
DIM bBackToForm : bBackToForm   = FALSE
DIM oRecord, oFile, oXMLDoc, sLog, sLogLine, oWaypoints
DIM sSQL, sLongitude, sLatitude
DIM oLongitude, oLatitude

CALL Feedback.setDefaults()

'do we go back to the form'
IF ( INSTR(sAction,"_form") > 0 ) THEN
  sAction     = REPLACE(sAction,"_form","")
  bBackToForm = TRUE
END IF

'custom validation'
IF ( NOT Feedback.Has("Error") ) THEN
  'defferent validatio n for different actions
  SELECT CASE UCASE(sAction)
    CASE "CREATE","UPDATE"
      IF ( UCASE(sAction) = "CREATE" AND NOT Feedback.Has("Error") ) THEN
        SET oLongitude  = oConvertToDegree(NULL,"lon")
        SET oLatitude   = oConvertToDegree(NULL,"lat")

        oLongitude.ITEM("degrees")  = c_FormRequest("LongitudeDegrese")
        oLongitude.ITEM("minutes")  = c_FormRequest("LongitudeMinutes")
        oLongitude.ITEM("dir")      = c_FormRequest("LongitudeDir")
        oLatitude.ITEM("degrees")   = c_FormRequest("LatitudeDegrees")
        oLatitude.ITEM("minutes")   = c_FormRequest("LatitudeMinutes")
        oLatitude.ITEM("dir")       = c_FormRequest("LatitudeDir")

        IF ( _
              bHaveInfo(c_FormRequest("Title")) AND _
              bHaveInfo(iConvertToDecimal(oLongitude,"lon")) AND _
              bHaveInfo(iConvertToDecimal(oLatitude,"lat")) _
         ) THEN

         sLongitude  = iConvertToDecimal(oLongitude,"lon")
         sLatitude   = iConvertToDecimal(oLatitude,"lat")

         'IF ( LEN(iConvertToDecimal(oLongitude,"lon")) > 10 ) THEN sLongitude = LEFT(iConvertToDecimal(oLongitude,"lon"),10)
         'IF ( LEN(iConvertToDecimal(oLatitude,"lat")) > 9 ) THEN sLatitude = LEFT(iConvertToDecimal(oLatitude,"lat"),9)

          sSQL = "SELECT * "
          sSQL = sSQL & "FROM Waypoint "
          sSQl = sSQL & "WHERE "
          sSQL = sSQL & " (Title)='" & UCASE(c_FormRequest("Title")) & "' "
          sSQL = sSQL & "OR "
          sSQL = sSQL & " ( "
          sSQL = sSQL & "   Longitude = '" & sLongitude & "' AND Latitude = '" &  sLatitude & "' "
          sSQL = sSQL & " ) "

          SET oRecord   = c_Waypoint.run("FindBySQL",sSQL)

          IF ( NOT oRecord IS NOTHING ) THEN
            IF ( UCASE(oRecord.FieldValue("Title")) = UCASE(c_FormRequest("Title")) ) THEN
              CALL Feedback.addToFeedback("<p>The Waypoint title already exists in your system.</p>")
            ELSE
              CALL Feedback.addToFeedback("<p>The Waypoint Longitude and Latitude already exists in your system (" & oRecord.FieldValue("Title") & ").</p>")
            END IF

            CALL Feedback.setError()
            CALL Feedback.setFormData()

            CALL Feedback.showFeedback("member-waypoints-form.asp")
          END IF
        END IF
      END IF

      IF ( NOT NOT Feedback.Has("Error") ) THEN

      END IF
    END SELECT
END IF

SELECT CASE UCASE(sAction)
  CASE "CREATE"
    CALL Feedback.setFeedbackAction("Create")

    SET oRecord   = c_Waypoint.run("New",NULL)

    CALL setFromForm(sAction,oRecord)

    IF ( oRecord.run("SaveRecord",NULL) ) THEN
      CALL saveSounderFields(sAction,oRecord)

      CALL Feedback.setComplete()
      CALL Feedback.addToFeedback("<p>Succesfully created your new waypoint (" & oRecord.FieldValue("Title") & ").</p>")

      IF ( bBackToForm ) THEN
        CALL Feedback.showFeedback("member-waypoints-form.asp")
      ELSE
        CALL Feedback.showFeedback("member-waypoints.asp")
      END IF
    ELSE
      CALL Feedback.addToFeedback("<p>unable to save record.</p>")

      CALL Feedback.setError()
      CALL Feedback.setFormData()

      CALL Feedback.showFeedback("member-waypoints-form.asp")
    END IF

  CASE "UPDATE"
    SET oRecord   = c_Waypoint.run("FindByID",c_FormRequest("Id"))

    IF ( NOT oRecord IS NOTHING  ) THEN
      CALL setFromForm(sAction,oRecord)

      IF ( oRecord.run("SaveRecord",NULL) ) THEN
        CALL saveSounderFields(sAction,oRecord)

        IF ( bBackToForm ) THEN
          CALL Feedback.showFeedback("member-waypoints-form.asp?id=" & oRecord.FieldValue("ID"))
        ELSE
          CALL Feedback.showFeedback("member-waypoints.asp")
        END IF
      ELSE
        CALL Feedback.addToFeedback("<p>unable to save record.</p>")

      END IF
    ELSE
      CALL Feedback.addToFeedback("<p>unable to find record.</p>")
    END IF

    CALL Feedback.setError()
    CALL Feedback.showFeedback("member-waypoints-form.asp")

  CASE "IMPORT"
    CALL Feedback.setFeedbackAction("Import")

    'Loop through all of the files in our upload object
    FOR EACH oFile IN c_FormFiles
    	'import to waypoint import directory
    	oFile.SAVEAS SERVER.MAPPATH(APPLICATION("WaypointsUpload")) & "\" & sFileFriendly(oFile.FileName)
    NEXT

    IF ( NOT c_FormFiles("ImportFile") IS NOTHING ) THEN
      SET oFile = c_FormFiles("ImportFile")

      'add to log'
      sLog = sLog & "<p>File succesfully uploaded and found.</p>"

      SELECT CASE CSTR(c_FormRequest("SounderID"))
        CASE "1" 'Raymarine Dragonfly
          SET oXMLDoc   = SERVER.CREATEOBJECT("MSXML2.DOMDocument.3.0")
          oXMLDoc.ASYNC = FALSE
          sLog          = ""
          sLogLine      = ""

          IF (  oXMLDoc.LOAD(oFile.PATH) ) THEN
            DIM oWpt, oWptExt
            DIM sName, sInternalID

            'add to log'
            sLog = sLog & "<p>XML Loaded succesfully.</p>"

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
                sSQL = sSQL & "         UCASE(Title)='" & UCASE(sName) & "' AND WaypointSounderData.[Value]='' "
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
            NEXT

            CALL Feedback.setComplete()
            CALL Feedback.setFeedbackHeading("Import was Succesfully")
          ELSE
            'add to log'
            sLog = sLog & "<p>Import file was not an XML file.</p>"

            CALL Feedback.setError()
            CALL Feedback.setFeedbackHeading("Import Failed")
          END IF

      END SELECT
    ELSE
      'add to log'
      sLog = sLog & "<p>No file uploaded and found.</p>"

      CALL Feedback.setError()
      CALL Feedback.setFeedbackHeading("Import Failed")
    END IF

    CALL Feedback.addToFeedback(sLog)
    CALL Feedback.showFeedback("member-waypoints-feedback.asp")

  CASE "EXPORT"
    CALL Feedback.setFeedbackAction("Export")

    IF ( bHaveInfo(c_FormRequest("SounderID")) ) THEN
      SELECT CASE CSTR(c_FormRequest("SounderID"))
        CASE "1" 'Raymarine Dragonfly
          SET oWaypoints  = c_Waypoint.run("FindByMember",Member.FieldValue("ID"))

          CALL createRaymarineDragonflyExport(oWaypoints,c_FormRequest("SounderID"))
          CALL Feedback.setComplete()
          CALL Feedback.setFeedbackHeading("Export was Succesfully")

        CASE ELSE
          'add to log'
          sLog = sLog & "<p>Selected Sounder not recognised.</p>"

          CALL Feedback.setError()
          CALL Feedback.setFeedbackHeading("Export Failed")

      END SELECT
    ELSE
      'add to log'
      sLog = sLog & "<p>No Sounder was selected.</p>"

      CALL Feedback.setError()
      CALL Feedback.setFeedbackHeading("Export Failed")
    END IF

    CALL Feedback.addToFeedback(sLog)

    IF ( Feedback.Has("Error") ) THEN
      CALL Feedback.showFeedback("member-waypoints-feedback.asp")
    END IF

  CASE "UNKNOWN"
    RESPONSE.WRITE "UNKNOWN<br />"

END SELECT
%>
