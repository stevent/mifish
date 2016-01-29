<%
'Ensure a new copy of the page every time
RESPONSE.EXPIRES	= -1
RESPONSE.BUFFER	= TRUE

'***************************************
' Copyright (c) 2001-2008 Webfirm
' http://www.webfirm.com.au
' Ph: 1300 304 779
'
' File:				    class.klass.asp
' Author:			    Steven Taddei
' Date Created:   18/06/2012
' Modified By:    Steven Taddei
' Date Modified:  02/07/2012
'
' Purpose:
'	This a class used for extending Base 
' and any other children classes.
'
' EXAMPLE OF USE: 
'   DIM Person : SET Person = Klass.new("Person",oSetParams(ARRAY("FirstName[]","Surname[]","Position[Developer]","DOB")))
'
'   Person.SetVal("FirstName")  = "Joe"
'   Person.SetVal("Surname")    = "Jackson"
'
'   FUNCTION Person_full_name(this,arrParams)
'     Person_full_name = this("FirstName") & " " & this("Surname")
'   END FUNCTION : Person.createMethod("full_name")
'
'   DIM Developer : SET Developer = Klass.extend(Person,"Developer",oSetParams(ARRAY("FirstName[Jake]","Surname[Smith]","Position[Rails Developer]","DOB")))
'
' FUNCTIONS
'   extend(oSuper,sName,oParams)
'     - extends a selected class object to 
'       a new object, oSuper is the elected object
'       sName is the New Class Name, oParams
'       is specific params of the new object
'
'   new(sName,oParams)
'     - creates a new class based on the 
'       base class. sname is the New Class
'       Name, oParams is specific params
'       of the new object
'***************************************

DIM Klass : SET Klass = NEW c_Klass

CLASS c_Klass
  '--------------------------------------------------------
  ' CLASS PUBLIC FUNCTIONS AND SUB ROUTINES
  '--------------------------------------------------------

	PUBLIC FUNCTION extend(BYREF oSuper,sName,oParams)
    DIM oNew : SET oNew = oSuper.new("")

    oNew.ClassName  = sName
    oNew.Super      = oSuper.ClassName

    CALL setInitialValues(oNew,oParams)
    CALL runInitializer(oNew)

		SET extend = oNew
	END FUNCTION

	PUBLIC FUNCTION [new](sName,oParams)
    DIM oBase : SET oBase = NEW C_Base

    oBase.ClassName = sName
    SET oBase.super = NEW C_Base

    CALL setInitialValues(oBase,oParams)
    CALL runInitializer(oBase)

		SET [new] = oBase
	END FUNCTION

  '--------------------------------------------------------
  ' CLASS PRIVATE FUNCTIONS AND SUB ROUTINES
  '--------------------------------------------------------

  PRIVATE SUB setInitialValues(oClass,oParams)
    DIM oField

    IF ( bIsDictionary(oParams) ) THEN

      FOR EACH oField IN oParams
        SET oField = oParams(oField)

        oClass.CreateField(oField.Name) = oField
      NEXT
    END IF
  END SUB

  PRIVATE SUB runInitializer(oClass)
    DIM bRunInitializer : bRunInitializer = FALSE
    DIM sTypeName

    ON ERROR RESUME NEXT

    sTypeName = typename(getRef(oClass.ClassName & "_Initialize"))

    IF ( ERR.NUMBER = 0 ) THEN bRunInitializer = TRUE

    ON ERROR GOTO 0

    IF ( bRunInitializer ) THEN
      EXECUTE("CALL " & oClass.ClassName & "_Initialize(oClass)")
    END IF
  END SUB

  '--------------------------------------------------------
  ' ON INITIALIZE AND ON TERMINATE
  '--------------------------------------------------------

  PRIVATE SUB setDefaults()

  END SUB

  PRIVATE SUB Class_Initialize()
    'set deafults
    CALL setDefaults()
  END SUB

  PRIVATE SUB Class_Terminate()
    'set deafults
    CALL setDefaults()
  END SUB
END CLASS
%>