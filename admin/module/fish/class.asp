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
' Date Created:   15/01/2015
' Modified By:    Steven Taddei
' Date Modified:  15/01/2015
'
' Purpose:
'	This is the code to build the Fish
' module
'
'***************************************

DIM c_Fish : SET c_Fish = Klass.extend(c_Module,"c_Fish",oSetParams(ARRAY("Table[Fish]","PrimaryKey[Id]","TablePlural[Fishes]")))

'--------------------------------------------------------
' SET UP CLASS VARIABLES
'--------------------------------------------------------

c_Fish.SetVar("TreeStructure")  = FALSE
c_Fish.SetVar("Table")          = "Fish"
c_Fish.SetVar("PrimaryKey")     = "ID"
c_Fish.SetVar("TablePlural")    = "Fishes"

'-------------------------------------------------------------------------------
' Purpose:  finds all the waypoints for a specific member
' Inputs:	  this    - the current Table object
'           oParams - any params we need to pass through for the function. Its
'                     either an ID or a dictionary params object
' Returns:	A specific list of records
'-------------------------------------------------------------------------------
FUNCTION c_Fish_FindByWaypoint(this,oParams)
  DIM iCount        : iCount        = 0
  DIM iRecordCount  : iRecordCount  = 0
  DIM sSQL, sDictKey
  DIM rsRecordSet
  DIM oNew, oTemp, oField, oWaypoints

  'set up sql
  sSQL = "SELECT " & this.Var("Table") & ".*"
  sSQL = sSQL & " FROM " & this.Var("Table") & ", WaypointCatch"
  sSQL = sSQL & " WHERE " & this.Var("Table") & ".ID=WaypointCatch.FishID"

  IF ( bIsDictionary(oParams) ) THEN
    IF ( oParams.EXISTS("conditions") ) THEN
      IF ( ISARRAY(oParams.ITEM("conditions")) ) THEN sSQL = sSQL & " AND " & setConditions(oParams.ITEM("conditions"))
    END IF

    IF ( oParams.EXISTS("order") ) THEN
      IF ( ISARRAY(oParams.ITEM("order")) ) THEN sSQL = sSQL & " ORDER BY " & setConditions(oParams.ITEM("order"))
    END IF
  ELSE
    sSQL = sSQL & " AND MemberID=" & iReturnInt(oParams)
    sSQL = sSQL & " ORDER BY Name"
  END IF

  'return new Object
  SET c_Fish_FindByWaypoint = this.run("Find",sSQL)
END FUNCTION : CALL c_Fish.createMethod("FindByWaypoint",TRUE)
%>
