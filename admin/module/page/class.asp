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
' Date Created:   02/07/2012
' Modified By:    Steven Taddei
' Date Modified:  02/07/2012
'
' Purpose:
'	This is the code to build the page
' module
'
'***************************************

DIM c_Page : SET c_Page = Klass.extend(c_Module,"c_Page",oSetParams(ARRAY("Table[Page]","PrimaryKey[Id]","TablePlural[Pages]")))

'--------------------------------------------------------
' SET UP CLASS VARIABLES
'--------------------------------------------------------

c_Page.SetVar("TreeStructure")  = FALSE
c_Page.SetVar("Table")          = "Page"
c_Page.SetVar("PrimaryKey")     = "ID"
c_Page.SetVar("TablePlural")    = "Pages"

'--------------------------------------------------------
' c_Page: KLASS PUBLIC FUNCTIONS AND SUB ROUTINES
'--------------------------------------------------------

'-------------------------------------------------------------------------------
' Purpose:  Returs the Heading
' Inputs:	  this    - the current Table object
'           iID     - the ID we are finding
' Returns:	The H1 Value (either the Heading field or default Name)
'-------------------------------------------------------------------------------
FUNCTION c_Page_getHeading(this,bWriteOut)
  DIM sTemp : sTemp = this.FieldValue("Name")

  IF ( bHaveInfo( this.FieldValue("Heading")) ) THEN sTemp = this.FieldValue("Heading")

  IF ( bWriteOut ) THEN
    RESPONSE.WRITE sTemp
  ELSE
    c_Page_getHeading = sTemp
  END IF
END FUNCTION : CALL c_Page.createMethod("getHeading",FALSE)
%>
