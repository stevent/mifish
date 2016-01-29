<%
'Ensure a new copy of the page every time
RESPONSE.EXPIRES	= -1
RESPONSE.BUFFER	= TRUE

'***************************************
' Copyright (c) 2001-2008 Webfirm
' http://www.webfirm.com.au
' Ph: 1300 304 779
'
' File:				    db.module.asp
' Author:			    Steven Taddei
' Date Created:   02/07/2012
' Modified By:    Steven Taddei
' Date Modified:  15/01/2015
'
' Purpose:
'	This is the base Module Object where
' everything extends off
'
'***************************************

DIM c_Module : SET c_Module = Klass.new("c_Module",oSetParams(ARRAY("Table[Page]","PrimaryKey[ID]")))

'--------------------------------------------------------
' SET UP CLASS VARIABLES
'--------------------------------------------------------

c_Module.SetVar("Table")          = ""
c_Module.SetVar("PrimaryKey")     = ""
c_Module.SetVar("TreeStructure")  = ""
c_Module.SetVar("TablePlural")    = ""
c_Module.SetVar("Validation")     = SERVER.CREATEOBJECT("SCRIPTING.DICTIONARY")

'--------------------------------------------------------
' ON INITIALIZE
'--------------------------------------------------------

SUB c_Module_Initialize(this)
  DIM iCount : iCount = 0
  DIM rsRecordSet
  DIM oTemp, oField

  'crate db recordset
  SET rsRecordSet = createReadonlyRecordset("SELECT * FROM " & this("Table") & " WHERE " & this("PrimaryKey") & "=0")

  'loop through all the fields
  FOR EACH oField IN rsRecordset.RS.Fields
    'create a field object in temp variables
    SET oTemp = NEW C_Like_Field

    'add field properties and data
    CALL addField(rsRecordset,oTemp,iCount)

    'create new field
    this.CreateField(oTemp.Name) = oTemp

    'release temp object
    SET oTemp = NOTHING

    'incrememnt field index by 1
    iCount = iCount + 1
  NEXT

  'release recordset
  SET rsRecordSet = NOTHING
END SUB

'--------------------------------------------------------
' c_Module: KLASS PUBLIC FUNCTIONS AND SUB ROUTINES
'--------------------------------------------------------

'-------------------------------------------------------------------------------
' Purpose:  creates a new instance based on the curent model
' Inputs:	  this    - the current Table object
'           oParams - any params we need to pass through for the function
' Returns:	A new instance
'-------------------------------------------------------------------------------
FUNCTION c_Module_New(this,oParams)
  DIM iCount : iCount = 0
  DIM rsRecordSet
  DIM oNew, oTemp, oField
  DIM sSQL

  'crate new instance
  SET oNew = this.new("")

  'set up sql
  sSQL = "SELECT * FROM " & this.Var("Table") & " WHERE " & this.Var("PrimaryKey") & "=0"

  SET rsRecordSet = createReadonlyRecordset(sSQL)

  'loop through all the fields
  FOR EACH oField IN rsRecordset.RS.Fields
    'create a field object in temp variables
    SET oTemp = NEW C_Like_Field

    'add field properties and data
    CALL addField(rsRecordset,oTemp,iCount)

    'create new field
    oNew.CreateField(oTemp.Name) = oTemp

    'release temp object
    SET oTemp = NOTHING

    'incrememnt field index by 1
    iCount = iCount + 1
  NEXT

  'release recordset
  SET rsRecordSet = NOTHING

  'return new Object
  SET c_Module_New = oNew
END FUNCTION : CALL c_Module.createMethod("New",TRUE)

'-------------------------------------------------------------------------------
' Purpose:  Find a bunch of records by SQL
' Inputs:	  this    - the current Table object
'           sSQL    - The sql for the databae query
' Returns:	A specific record based on the SQL passed through
'-------------------------------------------------------------------------------
FUNCTION c_Module_Find(this,sSQL)
  DIM iCount : iCount = 0
  DIM rsRecordSet
  DIM iRecordCount
  DIM oNew, oTemp, oField, oItems
  DIM sDictKey

  SET rsRecordSet = createReadonlyRecordset(sSQL)

  SET oItems = SERVER.CREATEOBJECT("SCRIPTING.DICTIONARY")

  'make sure we have records
  IF ( NOT rsRecordSet.RS.EOF ) THEN
    'loop through recordset
    DO WHILE ( NOT rsRecordSet.RS.EOF )
      'increment record count by 1
      iRecordCount  = iRecordCount + 1
      icount        = 0

      'crate new instance
      SET oNew = this.new("")

      'loop through all the fields
      FOR EACH oField IN rsRecordset.RS.Fields
        'create a field object in temp variables
        SET oTemp = NEW C_Like_Field

        'add field properties and data
        CALL addField(rsRecordset,oTemp,iCount)

        'create new field
        oNew.CreateField(oTemp.Name) = oTemp

        sDictKey = oNew.FieldValue("ID")

        'release temp object
        SET oTemp = NOTHING

        'incrememnt field index by 1
        iCount = iCount + 1
      NEXT

      CALL addToDictionatry(oItems,oNew,CSTR(sDictKey))

      rsRecordSet.RS.MOVENEXT
    LOOP
  END IF

  'release recordset
  SET rsRecordSet = NOTHING

  'return new Object
  SET c_Module_Find = oItems
END FUNCTION : CALL c_Module.createMethod("Find",TRUE)

'-------------------------------------------------------------------------------
' Purpose:  Find a specific record by SQL
' Inputs:	  this    - the current Table object
'           sSQL    - The sql for the databae query
' Returns:	A specific record based on the SQL passed through
'-------------------------------------------------------------------------------
FUNCTION c_Module_FindBySQL(this,sSQL)
  DIM oNew, oItems

  SET oItems = this.run("Find",sSQL)

  IF ( oItems.COUNT > 0 ) THEN
    SET oNew = oItems.ITEMS()(0)
  ELSE
    SET oNew = NOTHING
  END IF

  'return new Object
  SET c_Module_FindBySQL = oNew
END FUNCTION : CALL c_Module.createMethod("FindBySQL",TRUE)

'-------------------------------------------------------------------------------
' Purpose:  finds a specific DB table row based on the primary key ID
' Inputs:	  this    - the current Table object
'           oParams - any params we need to pass through for the function. Its
'                     either an ID or a dictionary params object
' Returns:	A specific record based on the iID passed through
'-------------------------------------------------------------------------------
FUNCTION c_Module_FindByID(this,oParams)
  DIM sSQL
  DIM iID

  iID = oParams

  'set up sql
  sSQL = "SELECT * FROM " & this.Var("Table") & " WHERE " & this.Var("PrimaryKey") & "=" & iReturnNumber(iID)

  'return new Object
  SET c_Module_FindByID = this.run("FindBySQL",sSQL)
END FUNCTION : CALL c_Module.createMethod("FindByID",TRUE)

'-------------------------------------------------------------------------------
' Purpose:  finds all records in a specifc table
' Inputs:	  this        - the current Table object
'           oParams     - any params we need to pass through for the function. Its
'                         either an ID or a dictionary params object
' Returns:	A specific list of records
'-------------------------------------------------------------------------------
FUNCTION c_Module_FindAll(this,oParams)
  DIM sSQL

  IF ( bIsDictionary(oParams) ) THEN
    'set up sql
    sSQL = "SELECT * "
    sSQL = sSQL & "FROM " & this.Var("Table") & " "

    IF ( oParams.EXISTS("conditions") ) THEN
      IF ( ISARRAY(oParams.ITEM("conditions")) ) THEN sSQL = sSQL & " WHERE " & setConditions(oParams.ITEM("conditions"))
    END IF

    IF ( oParams.EXISTS("order") ) THEN
      IF ( ISARRAY(oParams.ITEM("order")) ) THEN sSQL = sSQL & " ORDER BY " & setConditions(oParams.ITEM("order"))
    END IF
  ELSE
    'set up sql
    sSQL = "SELECT * FROM " & this.Var("Table")
  END IF

  'return new Object
  SET c_Module_FindAll = this.run("Find",sSQL)
END FUNCTION : CALL c_Module.createMethod("FindAll",TRUE)

'-------------------------------------------------------------------------------
' Purpose:  Saves the current object to the database
' Inputs:	  this    - the current Table object
'           oParams - any params we need to pass through for the function. Its
'                     either an ID or a dictionary params object
' Returns:	wether or not it was saved
'-------------------------------------------------------------------------------
FUNCTION c_Module_SaveRecord(this,oParams)
  DIM iCount  : iCount  = 0
  DIM bSaved  : bSaved  = FALSE
  DIM sSQL
  DIM rsRecordSet
  DIM iID
  DIM oNew, oTemp, oField

  iID = this.FieldValue(this.Var("PrimaryKey"))

  'set up sql
  sSQL = "SELECT * FROM " & this.Var("Table") & " WHERE " & this.Var("PrimaryKey") & "=" & iReturnNumber(iID)

  SET rsRecordSet = createUpdateableRecordset(sSQL)

  'loop through all the fields
  FOR EACH oField IN rsRecordset.RS.Fields
    IF ( NOT UCASE(this.Var("PrimaryKey")) = UCASE(rsRecordSet.RS.FIELDS(iCount).NAME) ) THEN
      rsRecordSet.RS.FIELDS(iCount).VALUE = this.FieldValue(CSTR(rsRecordSet.RS.FIELDS(iCount).NAME))
    END IF

    'incrememnt field index by 1
    iCount = iCount + 1
  NEXT

  rsRecordSet.RS.UPDATE

  this.SetValue(this.Var("PrimaryKey")) = rsRecordSet.FieldValue(this.Var("PrimaryKey"))

  bSaved = TRUE

  'return if its saved
  c_Module_SaveRecord = bSaved
END FUNCTION : CALL c_Module.createMethod("SaveRecord",FALSE)
%>