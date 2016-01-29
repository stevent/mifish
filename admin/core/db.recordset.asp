<%
FUNCTION createReadonlyRecordset(sSQL)
  DIM oTemp

  SET oTemp = NEW C_DBRecordset
  oTemp.SQL = sSQL
  CALL oTemp.setReadOnlyRecordset()

  SET createReadOnlyRecordset = oTemp
END FUNCTION

FUNCTION createUpdateableRecordset(sSQL)
  DIM oTemp

  SET oTemp = NEW C_DBRecordset
  oTemp.SQL = sSQL
  CALL oTemp.setUpdateableRecordset()

  SET createUpdateableRecordset = oTemp
END FUNCTION

CLASS C_DBRecordset
  PUBLIC Controller, Key
  PRIVATE m_recordset, m_HaveRecords, m_Size
  PRIVATE m_connection
  PRIVATE m_SQL

  'get the recordset and its variables
  PUBLIC PROPERTY GET RS()
    SET RS = m_recordset
  END PROPERTY
  PUBLIC PROPERTY GET HaveRecords()
    HaveRecords = m_HaveRecords
  END PROPERTY
  PUBLIC PROPERTY GET Size()
    Size = m_Size
  END PROPERTY
  PUBLIC PROPERTY GET FieldValue(sName)
    FieldValue = m_recordset(sName)
  END PROPERTY

  'get and set the sql for the recordset
  PUBLIC PROPERTY GET SQL()
    SQL = m_SQL
  END PROPERTY
  PUBLIC PROPERTY LET SQL(data)
    m_SQL = data
  END PROPERTY

  PUBLIC PROPERTY LET SetValue(sName,sValue)
    m_recordset(sName) = sValue
  END PROPERTY

  PRIVATE SUB Class_Initialize()
    'create our recordset
    SET  m_recordset = SERVER.CREATEOBJECT("ADODB.RecordSet")

    'create connection
    SET m_connection = createConnection()

    'default variables
    m_HaveRecords   = FALSE
    m_Size          = 0
  END SUB

  PUBLIC SUB setReadOnlyRecordset()
    m_recordset.ACTIVECONNECTION		  = m_connection.conn
    m_recordset.CURSORTYPE				    = 3	'adOpenStatic
    m_recordset.CURSORLOCATION			  = 3	'adUseClient
    m_recordset.LOCKTYPE				      = 3	'adLockOptimistic
    m_recordset.OPEN m_SQL
    SET m_recordset.ACTIVECONNECTION	= NOTHING

    IF ( NOT m_recordset.EOF ) THEN
      m_HaveRecords = TRUE
      m_Size        = m_recordset.RECORDCOUNT
    END IF
  END SUB

  PUBLIC SUB setUpdateableRecordset()
    SET m_recordset = SERVER.CREATEOBJECT("ADODB.Recordset")
    m_recordset.OPEN m_SQL, m_connection.conn, 1, 3

    IF ( m_recordset.EOF ) THEN
      m_recordset.ADDNEW
    END IF
  END SUB

  PRIVATE SUB Class_Terminate()
    'close and release the recordset
    IF ( ISOBJECT(m_recordset) AND NOT m_recordset IS NOTHING ) THEN
      SET m_recordset = NOTHING
    END IF

    'release connection
    SET m_connection = NOTHING
  END SUB
END CLASS
%>
