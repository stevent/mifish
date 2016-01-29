<%
'Ensure a new copy of the page every time
RESPONSE.EXPIRES	= -1
RESPONSE.BUFFER	= TRUE

'***************************************
' Copyright (c) 2001-2008 Webfirm
' http://www.webfirm.com.au
' Ph: 1300 304 779
'
' File:				    class.asp
' Author:			    Steven Taddei
' Date Created:   13/01/2015
' Modified By:    Steven Taddei
' Date Modified:  13/01/2015
'
' Purpose:
'	This is the code to build the type
' module
'
'***************************************

DIM c_WaypointType : SET c_WaypointType = Klass.extend(c_Module,"c_WaypointType",oSetParams(ARRAY("Table[WaypointType]","PrimaryKey[Id]","TablePlural[WaypointTypes]")))

'--------------------------------------------------------
' SET UP CLASS VARIABLES
'--------------------------------------------------------

c_WaypointType.SetVar("TreeStructure")  = FALSE
c_WaypointType.SetVar("Table")          = "WaypointType"
c_WaypointType.SetVar("PrimaryKey")     = "ID"
c_WaypointType.SetVar("TablePlural")    = "WaypointTypes"
%>
