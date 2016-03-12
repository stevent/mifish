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
SUB setCatchFromForm(sAction,oRecord)
  oRecord.SetValue("MemberID")      = Member.FieldValue("ID")
  oRecord.SetValue("WaypointID")    = c_FormRequest("WaypointID")
  oRecord.SetValue("FishID")        = c_FormRequest("FishID")
  oRecord.SetValue("Title")         = c_FormRequest("Title")
  oRecord.SetValue("CatchDate")     = c_FormRequest("CatchDate")
  oRecord.SetValue("Location")      = c_FormRequest("Location")
  oRecord.SetValue("TimeOfDay")     = c_FormRequest("TimeOfDay")
  oRecord.SetValue("Length")        = c_FormRequest("Length")
  oRecord.SetValue("Weight")        = c_FormRequest("Weight")
  oRecord.SetValue("Notes")         = c_FormRequest("Notes")
END SUB
%>
