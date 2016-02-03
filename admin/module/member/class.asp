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

DIM c_Member  : SET c_Member  = Klass.extend(c_Module,"c_Member",oSetParams(ARRAY("Table[Member]","PrimaryKey[Id]","TablePlural[Members]")))

'--------------------------------------------------------
' SET UP CLASS VARIABLES
'--------------------------------------------------------

c_Member.SetVar("TreeStructure")    = FALSE
c_Member.SetVar("Table")            = "Member"
c_Member.SetVar("PrimaryKey")       = "ID"
c_Member.SetVar("TablePlural")      = "Members"
c_Member.SetVar("MemberID_Cookie")  = "MID_MIFISH"
c_Member.SetVar("LoginURL")         = APPLICATION("SiteURL") & "/member-login.asp"
c_Member.SetVar("UsernameField")    = "Email"
c_Member.SetVar("PasswordField")    = "PasswordHash"

'--------------------------------------------------------
' c_Page: KLASS PUBLIC FUNCTIONS AND SUB ROUTINES
'--------------------------------------------------------

'-------------------------------------------------------------------------------
' Purpose:  Checks to see is a member is logged in
' Inputs:	  this    - the current Table object
'           oParams - params
' Returns:	TRUE|FALSE
'-------------------------------------------------------------------------------
FUNCTION c_Member_isLoggedIn(this,oParams)
  IF ( bHaveInfo(REQUEST.COOKIES(this.Var("MemberID_Cookie"))) ) THEN
    c_Member_isLoggedIn = TRUE
  ELSE
    c_Member_isLoggedIn = FALSE
  END IF
END FUNCTION : CALL c_Member.createMethod("isLoggedIn",FALSE)

'-------------------------------------------------------------------------------
' Purpose:  Logs out the currently logged in member
' Inputs:	  this    - the current Table object
'           oParams - any params we need to pass through for the function. Its
'                     either an ID or a dictionary params object
'-------------------------------------------------------------------------------
SUB c_Member_LogOut(this,oParams)
  RESPONSE.COOKIES(this.Var("MemberID_Cookie")).EXPIRES	= DATEADD("h",-1,DATE & " " & TIME)
END SUB : CALL c_Member.createMethod("LogOut",FALSE)

'-------------------------------------------------------------------------------
' Purpose:  Logs out the currently logged in member
' Inputs:	  this    - the current Table object
'           oParams - any params we need to pass through for the function. Its
'                     either an ID or a dictionary params object
'-------------------------------------------------------------------------------
SUB c_Member_SecurePage(this,oParams)
  IF ( NOT this.run("isLoggedIn",NULL) ) THEN
    RESPONSE.REDIRECT this.Var("LoginURL") & "?alert=" & SERVER.URLENCODE("you must log in")
  END IF
END SUB : CALL c_Member.createMethod("SecurePage",FALSE)

'-------------------------------------------------------------------------------
' Purpose:  Logs in a member
' Inputs:	  this    - the current Table object
'           oParams - any params we need to pass through for the function. Its
'                     either an ID or a dictionary params object
'-------------------------------------------------------------------------------
SUB c_Member_Login(this,oParams)
  DIM sUsername : sUsername = c_FormRequest("username")
  DIM sPassword : sPassword = c_FormRequest("password")
  DIM sSQL
  DIM rsRecordSet

  IF ( bHaveInfo(sUsername) AND bHaveInfo(sPassword) ) THEN
    'set up sql
    sSQL = "SELECT * "
    sSQL = sSQL & "FROM " & this.Var("Table") & " "
    sSQL = sSQL & "WHERE " & this.Var("UsernameField") & "='" & sUsername & "' "
    sSQL = sSQL & "AND " & this.Var("PasswordField") & "='" & sPassword & "' "

    SET rsRecordSet = createReadonlyRecordset(sSQL)

    'make sure we have records
    IF ( NOT rsRecordSet.RS.EOF ) THEN
      RESPONSE.COOKIES(this.Var("MemberID_Cookie")) = CSTR(rsRecordSet.RS.FIELDS.ITEM("ID"))
    END IF

    'release recordset
    SET rsRecordSet = NOTHING
  END IF

END SUB : CALL c_Member.createMethod("Login",FALSE)

'-------------------------------------------------------------------------------
' Purpose:  Grabs the currently logged in member
' Inputs:	  this    - the current Table object
'           oParams - any params we need to pass through for the function. Its
'                     either an ID or a dictionary params object
' Returns:	a Member Object (either full of info or empty)
'-------------------------------------------------------------------------------
FUNCTION c_Member_GrabLoggedInMember(this,oParams)
  DIM iID     : iID     = iReturnInt(REQUEST.COOKIES(this.Var("MemberID_Cookie")))
  DIM sSQL
  DIM oMember

  'set up sql
  sSQL = "SELECT * FROM " & this.Var("Table") & " WHERE " & this.Var("PrimaryKey") & "=" & iReturnInt(iID)

  SET oMember = this.run("FindBySQL",sSQL)

  IF ( oMember IS NOTHING ) THEN SET oMember = this.new("")

  'return new Object
  SET c_Member_GrabLoggedInMember = oMember
END FUNCTION : CALL c_Member.createMethod("GrabLoggedInMember",TRUE)

'--------------------------------------------------------
' DEFAULT INSTANCES
'--------------------------------------------------------
DIM Member : SET Member = c_Member.run("GrabLoggedInMember",NULL)
%>
