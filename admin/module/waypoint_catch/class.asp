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

DIM c_WaypointCatch : SET c_WaypointCatch = Klass.extend(c_Module,"c_WaypointCatch",oSetParams(ARRAY("Table[WaypointCatch]","PrimaryKey[Id]","TablePlural[WaypointCatches]")))

'--------------------------------------------------------
' SET UP CLASS VARIABLES
'--------------------------------------------------------

c_WaypointCatch.SetVar("TreeStructure")  = FALSE
c_WaypointCatch.SetVar("Table")          = "WaypointCatch"
c_WaypointCatch.SetVar("PrimaryKey")     = "ID"
c_WaypointCatch.SetVar("TablePlural")    = "WaypointCatches"
%>
