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

c_Member.SetVar("TreeStructure")        = FALSE
c_Member.SetVar("Table")                = "Member"
c_Member.SetVar("LoginsTable")          = "MemberLogin"
c_Member.SetVar("PrimaryKey")           = "ID"
c_Member.SetVar("TablePlural")          = "Members"
c_Member.SetVar("Login_Cookie")         = "LID_MIFISH"
c_Member.SetVar("Account_Cookie")       = "AID_MIFISH"
c_Member.SetVar("MemberTable_Cookie")   = "MTAB_MIFISH"
c_Member.SetVar("LoginURL")             = APPLICATION("SiteURL") & "/member-login.asp"
c_Member.SetVar("UsernameField")        = "Email"
c_Member.SetVar("PasswordField")        = "PasswordHash"
c_Member.SetVar("MasterAccount")        = FALSE
c_Member.SetVar("EditAccount")          = FALSE

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
  IF ( bHaveInfo(REQUEST.COOKIES(this.Var("Login_Cookie"))) ) THEN
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
  RESPONSE.COOKIES(this.Var("Login_Cookie")).EXPIRES	       = DATEADD("h",-1,DATE & " " & TIME)
  RESPONSE.COOKIES(this.Var("Account_Cookie")).EXPIRES	     = DATEADD("h",-1,DATE & " " & TIME)
  RESPONSE.COOKIES(this.Var("MemberTable_Cookie")).EXPIRES	 = DATEADD("h",-1,DATE & " " & TIME)
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
  DIM sUsername     : sUsername = c_FormRequest("username")
  DIM sPassword     : sPassword = c_FormRequest("password")
  DIM sPasswordHash
  DIM sSQL
  DIM rsRecordSet
  DIM sTableName, iID, iMemberID

  IF ( bHaveInfo(sUsername) AND bHaveInfo(sPassword) ) THEN
    'create hash of attempted password
    sPasswordHash = sha256(sPassword)

    'set up sql
    sSQL = "SELECT *, 0 AS MemberID "
    sSQL = sSQL & "FROM " & this.Var("Table") & " "
    sSQL = sSQL & "WHERE "
    sSQL = sSQL & " ("
    sSQl = sSQL & "   " & this.Var("UsernameField") & "='" & sUsername & "' "
    sSQL = sSQL & " AND "
    sSQL = sSQL & "   " & this.Var("PasswordField") & "='" & sPasswordHash & "' "
    sSQl = sSQL & " ) "

    SET rsRecordSet = createReadonlyRecordset(sSQL)

    'make sure we have records
    IF ( rsRecordSet.RS.EOF ) THEN
      sSQL = "SELECT * "
      sSQL = sSQL & "FROM " & this.Var("LoginsTable") & " "
      sSQL = sSQL & "WHERE "
      sSQL = sSQL & " ("
      sSQl = sSQL & "  " & this.Var("UsernameField") & "='" & sUsername & "' "
      sSQL = sSQL & " AND "
      sSQL = sSQL & "   " & this.Var("PasswordField") & "='" & sPasswordHash & "' "
      sSQl = sSQL & " ) "

      SET rsRecordSet = createReadonlyRecordset(sSQL)
    END IF

    'make sure we have records
    IF ( NOT rsRecordSet.RS.EOF ) THEN
      sTableName                      = "Member"
      iID                             = rsRecordSet.RS.FIELDS.ITEM("ID")
      iMemberID                       = iReturnInt(rsRecordSet.RS.FIELDS.ITEM("MemberID"))

      IF ( iMemberID = 0 ) THEN
        iMemberID   = iID
      ELSE
        sTableName  = "MemberLogin"
      END IF

      RESPONSE.COOKIES(this.Var("Login_Cookie"))       = CSTR(iID)
      RESPONSE.COOKIES(this.Var("Account_Cookie"))     = CSTR(iMemberID)
      RESPONSE.COOKIES(this.Var("MemberTable_Cookie")) = CSTR(sTableName)

      SET Member = this.run("GrabLoggedInMember",NULL)
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
  DIM iID     : iID         = iReturnInt(REQUEST.COOKIES(this.Var("Login_Cookie")))
  DIM sTable  : sTable      = sReturnDefaultString(REQUEST.COOKIES(this.Var("MemberTable_Cookie")), "Member")
  DIM sSQL    : sSQL        = ""
  DIM oMember : SET oMember = NOTHING

  IF ( bIsDictionary(oParams) ) THEN
    IF ( oParams.EXISTS("id") ) THEN
      iID = iReturnInt(oParams.ITEM("id"))
    END IF

    IF ( oParams.EXISTS("table") ) THEN
      sTable = sReturnDefaultString(oParams.ITEM("table"),"")
    END IF
  END IF

  'set up sql
  SELECT CASE UCASE (sTable)
    CASE "MEMBERLOGIN"
      sSQL = "SELECT "
      sSQL = sSQL & " " & this.Var("Table") & ".ID, "
      sSQL = sSQL & " " & this.Var("LoginsTable") & ".Email, "
      sSQL = sSQL & " " & this.Var("LoginsTable") & ".FirstName, "
      sSQL = sSQL & " " & this.Var("LoginsTable") & ".Surname "
      sSQL = sSQL & "FROM "
      sSQL = sSQL & " " & this.Var("Table") & ", " & this.Var("LoginsTable") & " "
      sSQL = sSQL & "WHERE "
      sSQL = sSQL & " " & this.Var("Table") & ".ID = " & this.Var("LoginsTable") & ".MemberID "
      sSQL = sSQL & "AND "
      sSQL = sSQL & " " & this.Var("LoginsTable") & ".ID=" & iReturnInt(iID) & " "

    CASE "MEMBER"
      sSQL = "SELECT * FROM " & this.Var("Table") & " WHERE " & this.Var("PrimaryKey") & "=" & iReturnInt(iID)

  END SELECT

  IF ( bHaveInfo(sSQL) ) THEN SET oMember = this.run("FindBySQL",sSQL)

  IF ( oMember IS NOTHING ) THEN SET oMember = c_Member.run("New",NULL)

  'return new Object
  SET c_Member_GrabLoggedInMember = oMember
END FUNCTION : CALL c_Member.createMethod("GrabLoggedInMember",TRUE)

'--------------------------------------------------------
' DEFAULT INSTANCES
'--------------------------------------------------------
DIM Member : SET Member = c_Member.run("GrabLoggedInMember",NULL)
%>
