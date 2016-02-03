<%
'Ensure a new copy of the page every time
RESPONSE.EXPIRES	= -1
RESPONSE.BUFFER	= TRUE

'***************************************
' Copyright (c) 2001-2008 Webfirm
' http://www.webfirm.com.au
' Ph: 1300 304 779
'
' File:				    class.field.asp
' Author:			    Steven Taddei
' Date Created:   18/06/2012
' Modified By:    Steven Taddei
' Date Modified:  02/07/2012
'
' Purpose:
'	Class for the field object
'***************************************

'C_Like_Field
'-------------------------------------------------------------------------------
' Purpose:    Creates an object wich holds specific information for each field
' PROPERTIES: Name          - Name of the field
'             Value         - Value of the field
'             FieldType    - Type of field
'             PrimaryKey   - Is the field a primary key
'             Unique        - Does the field need to be Unique
'             MaxSize      - The maximum size of the field
'             Required      - Is the field Required
'
' EXAMPLE OF USE: 
'   DIM oField = NEW C_Like_Field
'
'   oField.Value ="Hi"
'
'   RESPONSE.RITE oField.Value
'
' GETTERS
'   label()
'     - grabs a form label based on the field name
'
'   MustBeNumeric()
'     - Checks if the field must be numeric (based on the field type)
'-------------------------------------------------------------------------------
CLASS C_Like_Field
  PRIVATE m_Name
  PRIVATE m_Value
  PRIVATE m_FieldType
  PRIVATE m_RequiredInDB
  PRIVATE m_Required
  PRIVATE m_PrimaryKey
  PRIVATE m_MaxSize
  PRIVATE m_Unique
  PUBLIC Methods,Super, ClassName


  '--------------------------------------------------------
  ' CLASS GETTERS - Returns Values
  '--------------------------------------------------------

  PUBLIC PROPERTY GET label() : label = NameFromField(m_Name) : END PROPERTY
  PUBLIC PROPERTY GET MustBeNumeric() : MustBeNumeric = ( m_FieldType = "int" OR m_FieldType = "float" OR m_FieldType = "currency" ) : END PROPERTY
  PUBLIC PROPERTY GET Name() : Name = m_Name : END PROPERTY
  PUBLIC PROPERTY GET Value() : Value = m_Value : END PROPERTY
  PUBLIC PROPERTY GET FieldType() : FieldType = m_FieldType : END PROPERTY
  PUBLIC PROPERTY GET PrimaryKey() : PrimaryKey = m_PrimaryKey : END PROPERTY
  PUBLIC PROPERTY GET Unique() : Unique = m_Unique : END PROPERTY
  PUBLIC PROPERTY GET MaxSize() : MaxSize = m_MaxSize : END PROPERTY
  PUBLIC PROPERTY GET Required() : Required = m_Required : END PROPERTY


  '--------------------------------------------------------
  ' CLASS SETTERS - Sets Class Values
  '--------------------------------------------------------

  PUBLIC PROPERTY LET Value(data) : m_Value = sReturnNull(data) : END PROPERTY
  PUBLIC PROPERTY LET Name(data) : m_Name = sReturnNull(data) : END PROPERTY
  PUBLIC PROPERTY LET FieldType(data) : m_FieldType = sReturnNull(data) : END PROPERTY
  PUBLIC PROPERTY LET PrimaryKey(data) : m_PrimaryKey = bReturnBoolean(data) : END PROPERTY
  PUBLIC PROPERTY LET Unique(data) : m_Unique = bReturnBoolean(data) : END PROPERTY
  PUBLIC PROPERTY LET MaxSize(data) : m_MaxSize = iReturnInt(data) : END PROPERTY

  PUBLIC PROPERTY LET RequiredInDB(data)
    m_RequiredInDB  = bReturnBoolean(data)
    m_Required      = bReturnBoolean(data)
  END PROPERTY
  PUBLIC PROPERTY LET Required(data)
    IF ( NOT m_RequiredInDB ) THEN
      m_Required = bReturnBoolean(data)
    ELSE
      m_Required = m_RequiredInDB
    END IF
  END PROPERTY


  '--------------------------------------------------------
  ' CLASS PRIVATE FUNCTIONS AND SUB ROUTINES
  '--------------------------------------------------------

  PRIVATE SUB setDefaults()
    m_RequiredInDB  = FALSE
    m_Required      = FALSE
    m_PrimaryKey    = FALSE
    m_Unique        = FALSE
    m_MaxSize       = NULL
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
%>