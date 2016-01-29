<%
'Ensure a new copy of the page every time
RESPONSE.EXPIRES	= -1
RESPONSE.BUFFER	= TRUE

'***************************************
' Copyright (c) 2001-2008 Webfirm
' http://www.webfirm.com.au
' Ph: 1300 304 779
'
' File:				    core.request.asp
' Author:			    Steven Taddei
' Date Created:	  18/01/2015
' Last Modified:	18/01/2015
'
' Purpose:
'	Creates a request object in order to
' handle form posts and for gets.
'***************************************

DIM c_FormRequest, c_FormFiles, c_IsMultipart, c_ASP_Upload

IF ( NOT ISOBJECT(c_FormRequest) ) THEN
  CALL init_BuildRequest()
END IF

SUB init_BuildRequest()
  IF ( sGetFormEncType() = "MULTIPART/FORM-DATA" ) THEN
    SET c_ASP_Upload = SERVER.CREATEOBJECT("Persits.Upload.1")
    c_ASP_Upload.OVERWRITEFILES	= FALSE
    c_ASP_Upload.IGNORENOPOST		= TRUE

    'Save the files to memory temporarily
    c_ASP_Upload.SAVETOMEMORY

    SET c_FormRequest   = c_ASP_Upload.FORM
    SET c_FormFiles     = c_ASP_Upload.FILES
    c_IsMultipart       = TRUE
  ELSE
    SET c_FormRequest   = REQUEST.FORM
    c_FormFiles         = NULL
    c_IsMultipart       = FALSE
  END IF
END SUB

'-------------------------------------------------------------------------------
' Purpose:  To determine the encoding type of the page.
' Inputs:	None
' Returns:	The encoding type
'-------------------------------------------------------------------------------
FUNCTION sGetFormEncType()
  DIM iIndex
  DIM sContentType

  'Retrieve the content type field information
  sContentType = REQUEST.SERVERVARIABLES("CONTENT_TYPE")

  'Determine the position of the first semi colon in our content type
  iIndex = INSTR(sContentType, ";")

  'Check to see if we have found a semi colon in or content type string or not
  IF ( iIndex > 0 ) THEN
    'Retrieve the exact content type information
    sContentType = UCASE(TRIM(LEFT(sContentType, iIndex - 1)))
  ELSE
    'Retrieve the exact content type information
    sContentType = UCASE(TRIM(sContentType))
  END IF

  'Return the content type that we have found
  sGetFormEncType = sContentType
END FUNCTION
%>
