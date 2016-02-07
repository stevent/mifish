<%
'Ensure a new copy of the page every time
RESPONSE.EXPIRES	= -1
RESPONSE.BUFFER	= TRUE

'***************************************
' Copyright (c) 2001-2008 Webfirm
' http://www.webfirm.com.au
' Ph: 1300 304 779
'
' File:				    class.base.asp
' Author:			    Steven Taddei
' Date Created:   18/06/2012
' Modified By:    Steven Taddei
' Date Modified:  02/07/2012
'
' Purpose:
'	The base class that is inherited from
'
' GETTERS
'   FieldValue(sField)
'     - Gets the value for the specific Field
'   Var(sVariable)
'     - Allows you to set Class variables.
'       The values of there are kept when
'       extending the klass
'
'   FieldAttr(sField,sAttr)
'     - Gets the attribute value for the
'       specific Field
'
' SETTERS
'   CreateField(sField) = oField
'     - Sets a specific field with a field
'       object. field object is C_Like_Field
'       if field does not exist it is created
'   SetValue(sField) = Value
'     - Sets a specific field with a value.
'       if field does not exist it is created
'   SetVar(sVariable) = Value
'     - Sets a specific class variable. is the
'       variable does not exist it is created
'   SetFieldAttr(sFieldName,sAttr) = sValue
'     - Sets a value for a specific field
'       attribute
'   createMethod(sMethodName)
'     - Sets a link to runs a specific
'       method for the current object.
'       You write the function/sub outside
'       the class and the sMethodName is
'       the specifi name of it.
'       When you build the function/sub
'       the naming convention is
'       <ClassName>_<MethodName>.
'       EXAMPLE
'         FUNCTION Person_full_name(this,oParams)
'           Person_full_name = this("FirstName") & " " & this("Surname")
'         END FUNCTION : Person.createMethod("full_name")
'       oParams
'         oParams is a parameter object
'         created by the helper oSetParams
'         **see oSetParams in helpers.application.asp
'
' FUNCTIONS
'   createMethod(sMethodname)
'     - Creates a link to a specific
'       function so it can be fired
'       by the desired class
'   run(sMethodname,oFields)
'     - runs a specific method, finds it
'       by the method name and you are
'       able to pass through Fields.
'       the Fields are in the form of a
'       dictionary object
'         ' dict.ITEM(field_name) = value
'   new(oFields)
'     - creates a new instance of itself,
'       the instance is minus the data.
'       you are able to pass through Fields.,
'       wich is in the form of a
'       dictionary object
'         ' dict.ITEM(field_name) = value
'***************************************

CLASS C_Base
  PRIVATE Fields, ClassVariables
  PUBLIC Methods, Super, ClassName, MethodObjs

  '--------------------------------------------------------
  ' CLASS GETTERS - Returns Values
  '--------------------------------------------------------

  'FieldValue(sField)
  '   Grabs specific Field Value
  '   RESPONSE.WRITE C_Base.FieldValue("FieldName")
	PUBLIC DEFAULT FUNCTION FieldValue(sField)
    DIM oItem
    FOR EACH oItem IN Fields.ITEMS

    NEXT
    'lets make sure the field exists
    IF ( Fields.EXISTS(CSTR(sField)) ) THEN
      FieldValue = Fields(CSTR(sField)).Value
    ELSE
      'raise any errors
      PageError.RaiseError error_BaseFieldNotFound,"C_Base.FieldValue","The specific field does not exist (" & sField & ")"
    END IF
	END FUNCTION

  'FieldList()
  '   Grabs Field list in object
	PUBLIC FUNCTION FieldList()
    DIM oItem
    DIM sTemp

    FOR EACH oItem IN Fields.ITEMS
      IF ( bHaveInfo(sTemp) ) THEN sTemp = sTemp & ","
      sTemp = sTemp & oItem.NAME
    NEXT

    FieldList = sTemp
	END FUNCTION

  'Var(sVariable)
  '   Grabs specific Class Variable Value
  '   RESPONSE.WRITE C_Base.Var("VariableName")
	PUBLIC FUNCTION Var(sVariable)
    'lets make sure the field exists
    IF ( ClassVariables.EXISTS(CSTR(sVariable)) ) THEN
      Var = ClassVariables.ITEM(CSTR(sVariable))
    ELSE
      'raise any errors
      PageError.RaiseError error_BaseClassVariableNotFound,"C_Base.Var","The specific class variable does not exist (" & sVariable & ")"
    END IF
	END FUNCTION

  'FieldAttr(sField,Attr)
  '   Grabs specific Field Attribute
  '   RESPONSE.WRITE C_Base.FieldAttr("FieldName","MaxSize")
	PUBLIC FUNCTION FieldAttr(sField,Attr)
    'lets make sure the field exists
    IF ( Fields.EXISTS(CSTR(sField)) ) THEN
      EXECUTE("FieldAttr = Fields.ITEM(CSTR(sField))." & Attr)
    ELSE
      'raise any errors

      PageError.RaiseError error_BaseFieldNotFound,"C_Base.FieldAttr","The specific field does not exist (" & sField & ")"
    END IF
	END FUNCTION

  '--------------------------------------------------------
  ' CLASS SETTERS - Sets Class Values
  '--------------------------------------------------------

  'CreateField(sField,oField)
  '   Sets a specific field object to the slected class field
  '   C_Base.CreateField("FieldName") = oField
  PUBLIC PROPERTY LET CreateField(sField,oField)
    'add to Fields dictionary
    CALL addToDictionatry(Fields,oField,sField)
  END PROPERTY

  'SetValue(sField,Value)
  '   Sets the value for a specific field
  '   C_Base.SetValue("FieldName") = "Test Value"
  PUBLIC PROPERTY LET SetValue(sField,Value)
    DIM oField

    'Does the field exist? or do we create a new one?
    IF ( Fields.EXISTS(CSTR(sField)) ) THEN
      SET oField = Fields.ITEM(CSTR(sField))
    ELSE
      SET oField    = NEW C_Like_Field
      oField.Name   = Value
    END IF

    'set value
    oField.Value = Value

    'add to Fields dictionary
    CALL addToDictionatry(Fields,oField,oField.Name)
  END PROPERTY

  'SetVar(sVariable,Value)
  '   Sets a specific class variable
  '   C_Base.SetVar("FieldName") = "Test Value"
  PUBLIC PROPERTY LET SetVar(sVariable,Value)
    'add to Fields dictionary
    CALL addToDictionatry(ClassVariables,Value,sVariable)
  END PROPERTY

  'SetFieldAttr(sField,sAttr,sValue)
  '   Sets a value for a specific field attribute
  '   C_Base.SetFieldAttr("FieldName","Attr") = Value
	PUBLIC PROPERTY LET SetFieldAttr(sField,sAttr,sValue)
    'lets make sure the field exists
    IF ( Fields.EXISTS(CSTR(sField)) ) THEN
      EXECUTE("Fields.ITEM(CSTR(sField))." & sAttr & " = " & sValue)
    ELSE
      'raise any errors
      PageError.RaiseError error_BaseFieldNotFound,"C_Base.SetFieldAttr","The specific field does not exist (" & sField & ")"
    END IF
	END PROPERTY

  'createMethod(sField,sAttr,sValue)
  '   creates a specific link for a class method
  '   C_Base.createMethod("MethodName",FALSE)
	PUBLIC SUB createMethod(sMethodname,bIsObject)
    IF ( MethodObjs.EXISTS(CSTR(sMethodname)) OR Methods.EXISTS(CSTR(sMethodname)) ) THEN
      'raise any errors
      PageError.RaiseError error_BaseMethodAlreadyExists,"C_Base.createMethod","The specific method name already exists (" & sMethodname & ")"
    ELSE
      IF ( bIsObject ) THEN
        SET MethodObjs(CSTR(sMethodname)) = getRef(ClassName & "_" & sMethodname)
      ELSE
        SET Methods(CSTR(sMethodname)) = getRef(ClassName & "_" & sMethodname)
      END IF
    END IF
	END SUB

  '--------------------------------------------------------
  ' CLASS PUBLIC FUNCTIONS AND SUB ROUTINES
  '--------------------------------------------------------

  'run(sField,sAttr,sValue)
  '   run, runs a specific class method
  '   C_Base.run("MethodName",oFields)
  PUBLIC FUNCTION run(sMethodname,oFields)
    IF ( MethodObjs.EXISTS(CSTR(sMethodname)) ) THEN
      SET run = MethodObjs(CSTR(sMethodname))(Me,oFields)
    ELSEIF ( Methods.EXISTS(CSTR(sMethodname)) ) THEN
      run = Methods(CSTR(sMethodname))(Me,oFields)
    ELSE
      'raise any errors
      PageError.RaiseError error_BaseMethodDoesNotExist,"C_Base.run","The specific method does not exist (" & sMethodname & ")"
    END IF
  END FUNCTION

  'new(sField,sAttr,sValue)
  '   Create a new instance of the current object
  '   C_Base.new(oFields)
  PUBLIC FUNCTION [new](oFields)
    DIM oBase : SET oBase = NEW C_Base
    DIM oKey, oField

    IF ( ISOBJECT(Super) ) THEN
      SET oBase.Super   = Super
    ELSE
      oBase.Super   = Super
    END IF

    oBase.ClassName   = ClassName

		FOR EACH oKey IN ClassVariables
      oBase.SetVar(CSTR(oKey)) = ClassVariables.ITEM(oKey)
		NEXT

		FOR EACH oKey IN Fields
      'craete field object
      SET oField    = NEW C_Like_Field

      oField.Value          = ""
      oField.Name           = Fields(CSTR(oKey)).Name
      oField.FieldType      = Fields(CSTR(oKey)).FieldType
      oField.PrimaryKey     = Fields(CSTR(oKey)).PrimaryKey
      oField.Unique         = Fields(CSTR(oKey)).Unique
      oField.MaxSize        = Fields(CSTR(oKey)).MaxSize

      oBase.CreateField(CSTR(oKey)) = oField
		NEXT

		FOR EACH oKey IN Methods
			IF ISOBJECT(Methods(oKey)) then
				SET oBase.Methods(CSTR(oKey)) = Methods(CSTR(oKey))
			ELSE
				oBase.Methods(CSTR(oKey)) = Methods(CSTR(oKey))
			END IF
		NEXT

		FOR EACH oKey IN MethodObjs
			IF ISOBJECT(MethodObjs(oKey)) then
				SET oBase.MethodObjs(CSTR(oKey)) = MethodObjs(CSTR(oKey))
			ELSE
				oBase.MethodObjs(CSTR(oKey)) = MethodObjs(CSTR(oKey))
			END IF
		NEXT

    CALL setInitialValues(oBase,oFields)

    SET [new] = oBase
  END FUNCTION

  '--------------------------------------------------------
  ' CLASS PRIVATE FUNCTIONS AND SUB ROUTINES
  '--------------------------------------------------------

  'setInitialValues(obj,arr_values)
  '   Sets initial variables from passed object
  PRIVATE SUB setInitialValues(oClass,oParams)
    DIM oField

    IF ( bIsDictionary(oParams) ) THEN

      FOR EACH oField IN oParams
        SET oField = oParams(oField)

        oClass.CreateField(oField.Name) = oField
      NEXT
    END IF
  END SUB

  PRIVATE SUB setDefaults()
    SET Fields          = SERVER.CREATEOBJECT("SCRIPTING.DICTIONARY")
    SET ClassVariables  = SERVER.CREATEOBJECT("SCRIPTING.DICTIONARY")
    SET Methods         = SERVER.CREATEOBJECT("SCRIPTING.DICTIONARY")
    SET MethodObjs      = SERVER.CREATEOBJECT("SCRIPTING.DICTIONARY")
    Super               = NULL
    ClassName           = "C_Base"
  END SUB

  '--------------------------------------------------------
  ' ON INITIALIZE AND ON TERMINATE
  '--------------------------------------------------------

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
