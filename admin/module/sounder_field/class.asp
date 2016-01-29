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
'	This is the code to build the sounder
' fields module
'
'***************************************

DIM c_SounderField : SET c_SounderField = Klass.extend(c_Module,"c_SounderField",oSetParams(ARRAY("Table[SounderField]","PrimaryKey[Id]","TablePlural[SounderFields]")))

'--------------------------------------------------------
' SET UP CLASS VARIABLES
'--------------------------------------------------------

c_SounderField.SetVar("TreeStructure")  = FALSE
c_SounderField.SetVar("Table")          = "SounderField"
c_SounderField.SetVar("PrimaryKey")     = "ID"
c_SounderField.SetVar("TablePlural")    = "SounderFields"

'-------------------------------------------------------------------------------
' Purpose:  finds all the waypoints for a specific member
' Inputs:	  this    - the current Table object
'           oParams - any params we need to pass through for the function. Its
'                     either an ID or a dictionary params object
' Returns:	A specific list of records
'-------------------------------------------------------------------------------
FUNCTION c_SounderField_FindBySounder(this,oParams)
  DIM iSounderID, iWaypointID
  DIM sSQL

  IF ( ISOBJECT(oParams) ) THEN
    IF ( oParams.EXISTS("iWaypointID") ) THEN iWaypointID = oParams.ITEM("iWaypointID").Value
    IF ( oParams.EXISTS("iSounderID") ) THEN iSounderID = oParams.ITEM("iSounderID").Value
  ELSE
    iSounderID = oParams
  END IF

  'set up sql
  sSQL = "SELECT *, "
  sSQL = sSQL & "(SELECT Value FROM WaypointSounderData WHERE WaypointSounderData.SounderFieldID=" & this.Var("Table") & ".ID AND WaypointID=" & iReturnNumber(iWaypointID) & ") AS FieldValue "
  sSQL = sSQL & "FROM " & this.Var("Table") & " "
  sSQL = sSQL & "WHERE " & this.Var("Table") & ".SounderID=" & iReturnNumber(iSounderID) & " "

  'return new Object
  SET c_SounderField_FindBySounder = this.run("Find",sSQL)
END FUNCTION : CALL c_SounderField.createMethod("FindBySounder",TRUE)

'-------------------------------------------------------------------------------
' Purpose:  Saves the current object to the database
' Inputs:	  this    - the current Table object
'           oParams - any params we need to pass through for the function. Its
'                     either an ID or a dictionary params object
' Returns:	wether or not it was saved
'-------------------------------------------------------------------------------
FUNCTION c_SounderField_SaveWaypointData(this,oParams)
  DIM iCount  : iCount  = 0
  DIM bValid  : bValid  = FALSE
  DIM bSaved  : bSaved  = FALSE
  DIM sSQL
  DIM rsRecordSet
  DIM iWaypointSounderData, iWaypointID, iSounderID, iSounderFieldID

  IF ( ISOBJECT(oParams) ) THEN
    IF ( oParams.EXISTS("iWaypointID") ) THEN iWaypointID = oParams.ITEM("iWaypointID").Value
    IF ( oParams.EXISTS("iSounderID") ) THEN iSounderID = oParams.ITEM("iSounderID").Value
    IF ( oParams.EXISTS("iSounderFieldID") ) THEN iSounderFieldID = oParams.ITEM("iSounderFieldID").Value

    'set up sql
    sSQL = "SELECT * "
    sSQL = sSQL & "FROM WaypointSounderData "
    sSQL = sSQL & "WHERE WaypointID=" & iReturnNumber(iWaypointID) & " "
    sSQL = sSQL & "AND SounderID=" & iReturnNumber(iSounderID) & " "
    sSQL = sSQL & "AND SounderFieldID=" & iReturnNumber(iSounderFieldID)
  ELSE
    iWaypointSounderData = oParams

    'set up sql
    sSQL = "SELECT * FROM WaypointSounderData WHERE ID=" & iReturnNumber(iWaypointSounderData)
  END IF

  SET rsRecordSet = createUpdateableRecordset(sSQL)

  IF ( iReturnNumber(iSounderID) > 0 ) THEN rsRecordSet.SetValue("SounderID") = iSounderID
  IF ( iReturnNumber(iWaypointID) > 0 ) THEN rsRecordSet.SetValue("WaypointID") = iWaypointID
  IF ( iReturnNumber(iSounderFieldID) > 0 ) THEN rsRecordSet.SetValue("SounderFieldID") = iSounderFieldID

  rsRecordSet.SetValue("Value") = this.FieldValue("FieldValue")

  IF ( iReturnNumber(rsRecordSet.FieldValue("SounderID")) > 0 AND iReturnNumber(rsRecordSet.FieldValue("WaypointID")) > 0 AND iReturnNumber(rsRecordSet.FieldValue("SounderFieldID")) > 0 ) THEN
    rsRecordSet.RS.UPDATE

    bSaved = TRUE
  ELSE
    'addToDictionatry'
  END IF

  'return if its saved
  c_SounderField_SaveWaypointData = bSaved
END FUNCTION : CALL c_SounderField.createMethod("SaveWaypointData",FALSE)
%>
