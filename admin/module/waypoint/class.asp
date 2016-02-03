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
'	This is the code to build the waypoint
' module
'
'***************************************

DIM c_Waypoint : SET c_Waypoint = Klass.extend(c_Module,"c_Waypoint",oSetParams(ARRAY("Table[Waypoint]","PrimaryKey[Id]","TablePlural[Waypoints]")))

'--------------------------------------------------------
' SET UP CLASS VARIABLES
'--------------------------------------------------------

c_Waypoint.SetVar("TreeStructure")  = FALSE
c_Waypoint.SetVar("Table")          = "Waypoint"
c_Waypoint.SetVar("PrimaryKey")     = "ID"
c_Waypoint.SetVar("TablePlural")    = "Waypoints"
c_Waypoint.SetVar("HasMany")        = "WayPointSounderData"

'-------------------------------------------------------------------------------
' Purpose:  finds all the waypoints for a specific member
' Inputs:	  this    - the current Table object
'           oParams - any params we need to pass through for the function. Its
'                     either an ID or a dictionary params object
' Returns:	A specific list of records
'-------------------------------------------------------------------------------
FUNCTION c_Waypoint_FindByMember(this,oParams)
  DIM sSQL
  DIm iMemberID

  IF ( bIsDictionary(oParams) ) THEN
    IF ( oParams.EXISTS("conditions") ) THEN
      IF ( ISARRAY(oParams.ITEM("conditions")) ) THEN sSQL = sSQL & " WHERE " & setConditions(oParams.ITEM("conditions"))
    END IF

    IF ( oParams.EXISTS("order") ) THEN
      IF ( ISARRAY(oParams.ITEM("order")) ) THEN sSQL = sSQL & " ORDER BY " & setConditions(oParams.ITEM("order"))
    END IF
  ELSE
    iMemberID       = iReturnInt(oParams)
    SET oParams     = oSetParams(ARRAY("conditions[]","order[]"))

    oParams.ITEM("conditions")  = ARRAY("MemberID = " & setParamNumber(1),iMemberID)
    oParams.ITEM("order")       = ARRAY("Title")
  END IF

  SET c_Waypoint_FindByMember = this.run("FindAll",oParams)
END FUNCTION : CALL c_Waypoint.createMethod("FindByMember",TRUE)
%>
