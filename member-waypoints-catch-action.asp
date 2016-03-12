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

SERVER.SCRIPTTIMEOUT = 1200

%>
<!--#include virtual="/admin/core/core.includes.asp"-->
<%
CALL Member.run("Login",NULL)
CALL Member.run("SecurePage",NULL)

DIM sAction     : sAction         = sReturnDefaultString(c_FormRequest("Action"),sReturnDefaultString(REQUEST.QUERYSTRING("Action"),"unknown"))
DIM oWaypoint   : SET oWaypoint   = c_Waypoint.run("FindByID",INT(sReturnDefaultString(c_FormRequest("WaypointID"),sReturnDefaultString(REQUEST.QUERYSTRING("WaypointID"),"0"))))
DIM bBackToForm : bBackToForm     = FALSE
DIM oRecord, oFile, oXMLDoc, sLog, sLogLine, oWaypoints
DIM sSQL, sLongitude, sLatitude
DIM oLongitude, oLatitude

CALL Feedback.setDefaults()

'do we go back to the form'
IF ( INSTR(sAction,"_form") > 0 ) THEN
  sAction     = REPLACE(sAction,"_form","")
  bBackToForm = TRUE
END IF

IF ( oWaypoint IS NOTHING ) THEN
  sAction         = "no waypoint"
END IF

'custom validation'
IF ( NOT Feedback.Has("Error") ) THEN
  'defferent validatio n for different actions
  SELECT CASE UCASE(sAction)
    CASE "CREATE","UPDATE"

  END SELECT
END IF

SELECT CASE UCASE(sAction)
  CASE "NO WAYPOINT"
    'add to log'
    CALL Feedback.setError()
    CALL Feedback.setFeedbackHeading("No Waypoint Found")
    CALL Feedback.addToFeedback("<p>No waypoint was found to add the catch to.</p>")
    CALL Feedback.showFeedback("member-waypoints-catch.asp?WaypointID=" & oWaypoint.FieldValue("ID"))

  CASE "CREATE"
    CALL Feedback.setFeedbackAction("Create")

    SET oRecord   = c_WaypointCatch.run("New",NULL)

    CALL setCatchFromForm(sAction,oRecord)

    IF ( oRecord.run("SaveRecord",NULL) ) THEN
      CALL Feedback.setComplete()
      CALL Feedback.addToFeedback("<p>Succesfully added your catch (" & oRecord.FieldValue("Title") & ").</p>")

      IF ( bBackToForm ) THEN
        CALL Feedback.showFeedback("member-waypoints-catch-form.asp?WaypointID=" & oWaypoint.FieldValue("ID"))
      ELSE
        CALL Feedback.showFeedback("member-waypoints-catch.asp?WaypointID=" & oWaypoint.FieldValue("ID"))
      END IF
    ELSE
      CALL Feedback.addToFeedback("<p>unable to save record.</p>")

      CALL Feedback.setError()
      CALL Feedback.setFormData()

      CALL Feedback.showFeedback("member-waypoints-catch-form.asp?WaypointID=" & oWaypoint.FieldValue("ID"))
    END IF

  CASE "UPDATE"
    CALL Feedback.setFeedbackAction("Update")

    SET oRecord   = c_WaypointCatch.run("FindByID",c_FormRequest("Id"))

    IF ( NOT oRecord IS NOTHING  ) THEN
      CALL setCatchFromForm(sAction,oRecord)

      IF ( oRecord.run("SaveRecord",NULL) ) THEN
        CALL Feedback.setComplete()
        CALL Feedback.addToFeedback("<p>Succesfully edited your catch (" & oRecord.FieldValue("Title") & ").</p>")

        IF ( bBackToForm ) THEN
          CALL Feedback.showFeedback("member-waypoints-catch-form.asp?id=" & oRecord.FieldValue("ID") & "&WaypointID=" & oWaypoint.FieldValue("ID"))
        ELSE
          CALL Feedback.showFeedback("member-waypoints-catch.asp?WaypointID=" & oWaypoint.FieldValue("ID"))
        END IF
      ELSE
        CALL Feedback.addToFeedback("<p>unable to save record.</p>")

      END IF
    ELSE
      CALL Feedback.addToFeedback("<p>unable to find record.</p>")
    END IF

    CALL Feedback.setError()
    CALL Feedback.showFeedback("member-waypoints-catch-form.asp?WaypointID=" & oWaypoint.FieldValue("ID"))

  CASE "DEL"
    CALL Feedback.setFeedbackAction("Delete")

    SET oRecord   = c_Waypoint.run("FindByID",REQUEST.QUERYSTRING("Id"))

    IF ( NOT oRecord IS NOTHING  ) THEN
      IF ( oRecord.run("DeleteRecord",NULL) ) THEN
        CALL Feedback.setComplete()
        CALL Feedback.addToFeedback("<p>Succesfully deleted your catch</p>")

        'create JSON data file
        CALL createJsonDataFile()

        CALL Feedback.showFeedback("member-waypoints-catch.asp?WaypointID=" & oWaypoint.FieldValue("ID"))
      ELSE
        CALL Feedback.addToFeedback("<p>unable to delete record.</p>")
      END IF
    ELSE
      CALL Feedback.addToFeedback("<p>unable to find record.</p>")
    END IF

    CALL Feedback.setError()
    CALL Feedback.showFeedback("member-waypoints-catch.asp?WaypointID=" & oWaypoint.FieldValue("ID"))

  CASE "UNKNOWN"
    'add to log'
    sLog = sLog & "<p>The selected action could not be found. It is unknown</p>"

    CALL Feedback.setError()
    CALL Feedback.setFeedbackHeading("Unknow Action")
    CALL Feedback.showFeedback("member-waypoints-catch.asp?WaypointID=" & oWaypoint.FieldValue("ID"))

END SELECT
%>
