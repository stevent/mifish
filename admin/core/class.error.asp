<%
'Ensure a new copy of the page every time
RESPONSE.EXPIRES	= -1
RESPONSE.BUFFER	= TRUE

'***************************************
' Copyright (c) 2001-2008 Webfirm
' http://www.webfirm.com.au
' Ph: 1300 304 779
'
' File:				    class.error.asp
' Author:			    Steven Taddei
' Date Created:   18/06/2012
' Modified By:    Steven Taddei
' Date Modified:  02/07/2012
'
' Purpose:
'	Class object for error catching/raising
'
' FUNCTIONS
'   CALL RaiseError(iNumber,sSection,sMessage)
'     -  Allows you to raise errors that 
'       are not general classic ASP errors
'***************************************

DIM PageError

'initialize error
CALL INIT_Error()

'INIT_Error
'-------------------------------------------------------------------------------
' Purpose:    Creates the Error object if it is not already created
'-------------------------------------------------------------------------------
SUB INIT_Error
  IF ( NOT ISOBJECT(PageError) ) THEN
    SET PageError = NEW C_Error
  END IF
END SUB

CLASS C_Error
  '--------------------------------------------------------
  ' CLASS PUBLIC FUNCTIONS AND SUB ROUTINES
  '--------------------------------------------------------

  PUBLIC FUNCTION RaiseError(iNumber,sSection,sMessage)
    RaiseError = EXECUTE("CALL ERR.Raise(iNumber, sSection, sMessage)")
  END FUNCTION

  '--------------------------------------------------------
  ' CLASS PRIVATE FUNCTIONS AND SUB ROUTINES
  '--------------------------------------------------------

  PRIVATE SUB setDefaults()

  END SUB


  '--------------------------------------------------------
  ' ON INITIALIZE AND ON TERMINATE
  '--------------------------------------------------------

  'on object initialization
  PRIVATE SUB Class_Initialize()
    CALL setDefaults()
  END SUB

  'on object termination
  PRIVATE SUB Class_Terminate()
    CALL setDefaults()
  END SUB
END CLASS


'--------------------------------------------------------
' CUSTOM ERROR NUMBERS
'--------------------------------------------------------
CONST error_BaseFieldNotFound           = 100
CONST error_BaseMethodAlreadyExists     = 101
CONST error_BaseMethodDoesNotExist      = 102
CONST error_BaseClassVariableNotFound   = 103
%>