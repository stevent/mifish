<%
DIM o_ConnProperties : SET o_ConnProperties = NEW C_ConnProperties

SUB giveBackConnection()
  CALL o_ConnProperties.hijack("")
END SUB

SUB hijackConnection(ConnString)
  CALL o_ConnProperties.hijack(ConnString)
END SUB

FUNCTION createConnection()
  DIM oTemp

  SET oTemp = NEW C_DBConn
  CALL oTemp.setConnection()

  SET createConnection = oTemp
END FUNCTION

CLASS C_DBConn
  PRIVATE m_conn

  PUBLIC PROPERTY GET Conn()
    SET Conn = m_conn
  END PROPERTY

  PUBLIC SUB setConnection()
    'close (if possible)
    IF ( m_conn.State <> 0 ) THEN m_conn.Close

    'set connection string
    m_conn.CONNECTIONSTRING  = o_ConnProperties.GrabConnString

    'open connection
    m_conn.OPEN
  END SUB

  PRIVATE SUB Class_Initialize()
    'create connection object
    SET m_conn = SERVER.CREATEOBJECT("ADODB.Connection")
  END SUB

  PRIVATE SUB Class_Terminate()
    'close and release the connection (if possible)
    IF ( m_conn.State <> 0 ) THEN m_conn.CLOSE

    'set connection to nothing
    SET m_conn = NOTHING
  END SUB
END CLASS

CLASS C_ConnProperties
  PRIVATE m_DefaultConnString, m_NewConnString

  PUBLIC PROPERTY GET GrabConnString
    IF ( bHaveInfo(m_NewConnString) ) THEN
      GrabConnString = m_NewConnString
    ELSE
      GrabConnString = m_DefaultConnString
    END IF
  END PROPERTY

  PUBLIC PROPERTY GET DefaultConnString()
    DefaultConnString = m_DefaultConnString
  END PROPERTY

  PUBLIC PROPERTY GET NewConnString()
    NewConnString = m_NewConnString
  END PROPERTY

  PUBLIC SUB hijack(data)
    m_NewConnString = data
  END SUB

  PRIVATE SUB Class_Initialize()
    '## What happens when the class is opened
    m_DefaultConnString = APPLICATION("ConString")
  END SUB

  PRIVATE SUB Class_Terminate()
    '## What happens when the class is closed
    m_DefaultConnString = ""
  END SUB
END CLASS
%>
