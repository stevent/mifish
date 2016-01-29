<%
'Ensure a new copy of the page every time
RESPONSE.EXPIRES	= -1
RESPONSE.BUFFER	= TRUE

'***************************************
' Copyright (c) 2001-2008 Webfirm
' http://www.webfirm.com.au
' Ph: 1300 304 779
'
' File:				    core.feedback.asp
' Author:			    Steven Taddei
' Date Created:	  18/01/2015
' Last Modified:	18/01/2015
'
' Purpose:
'	Object to handle site/system feedback
' messages
'***************************************

DIM Feedback

IF ( NOT ISOBJECT(Feedback) ) THEN
  CALL init_BuildFeedback()
END IF

SUB init_BuildFeedback()
  SET Feedback = NEW c_Feedback

  CALL Feedback.setFeedbackType(SESSION("FeedbackType_MIFISH"))
  CALL Feedback.addToFeedback(SESSION("Feedback_MIFISH"))
  CALL Feedback.setFeedbackPage(SESSION("FeedbackPage_MIFISH"))
  CALL Feedback.setFeedbackHeading(SESSION("FeedbackHeading_MIFISH"))
  CALL Feedback.setFeedbackAction(SESSION("FeedbackAction_MIFISH"))

  IF ( NOT UCASE(getCurrentPage()) = UCASE(Feedback.Page) ) then
    CALL Feedback.ClearSessions()
  END IF
END SUB

CLASS c_Feedback
  PRIVATE c_Type, c_Feedback, c_FeedbackPage, c_FeedbackHeading, c_FeedbackAction

  PUBLIC PROPERTY GET Page()
    Page = c_FeedbackPage
  END PROPERTY

  PUBLIC PROPERTY GET Action()
    Action = c_FeedbackAction
  END PROPERTY

  PUBLIC PROPERTY GET Content()
    Content = c_Feedback
  END PROPERTY

  PUBLIC PROPERTY GET FeedbackType()
    FeedbackType = c_Type
  END PROPERTY

  PUBLIC PROPERTY GET Heading()
    Heading = c_FeedbackHeading
  END PROPERTY

  PUBLIC FUNCTION Has(sType)
    DIM iTypeSelect : iTypeSelect = 0
    DIM bTemp       : bTemp       = FALSE

    SELECT CASE UCASE(sType)
      CASE "ERROR"
        iTypeSelect = 3

      CASE "WARNING"
        iTypeSelect = 2

      CASE "ALERT"
        iTypeSelect = 1

      CASE "COMPLETED"
        iTypeSelect = 4

    END SELECT

    IF ( iTypeSelect = c_Type ) THEN
      IF ( bHaveInfo(c_Feedback) ) THEN
        bTemp = TRUE
      END IF
    END IF

    Has = bTemp
  END FUNCTION

  PUBLIC SUB setFeedbackType(iType)
    'Alert        = 1'
    'Warning      = 2'
    'Error        = 3'
    'Complete     = 4'
    'information  = 5'

    IF ( iReturnNumber(iType) > 0 ) THEN
      SELECT CASE iReturnNumber(iType)
        CASE 1,2,3,4,5
          c_Type = iReturnNumber(iType)

      END SELECT
    END IF
  END SUB

  PUBLIC FUNCTION IsEmpty()
    IsEmpty = ( NOT bHaveInfo(c_Feedback) )
  END FUNCTION

  PUBLIC SUB setFeedbackPage(sText)
    c_FeedbackPage = sText
  END SUB

  PUBLIC SUB setFeedbackHeading(sText)
    c_FeedbackHeading = sText
  END SUB

  PUBLIC SUB setFeedbackAction(sText)
    c_FeedbackAction = sText
  END SUB

  PUBLIC SUB addToFeedback(sText)
    c_Feedback = c_Feedback & sText
  END SUB

  PUBLIC SUB showFeedback(sURL)
    IF ( bHaveInfo(sURL) ) THEN
      SESSION("FeedbackType_MIFISH")    = c_Type
      SESSION("Feedback_MIFISH")        = c_Feedback
      SESSION("FeedbackPage_MIFISH")    = sURL
      SESSION("FeedbackHeading_MIFISH") = c_FeedbackHeading
      SESSION("FeedbackAction_MIFISH")  = c_FeedbackAction

      RESPONSE.REDIRECT APPLICATION("SiteURL") & "/" & sURL
    END IF
  END SUB

  PUBLIC SUB setError()
    setFeedbackType(3)
  END SUB

  PUBLIC SUB setComplete()
    setFeedbackType(4)
  END SUB

  PUBLIC SUB setWarning()
    setFeedbackType(2)
  END SUB

  PUBLIC SUB setFormData()
    DIM oItem

    FOR EACH oItem IN c_FormRequest
      SESSION("FeedbackFormField_" & oItem & "_MIFISH") = c_FormRequest(oItem)
    NEXT
  END SUB

  PUBLIC FUNCTION getField(sName)
    getField = SESSION("FeedbackFormField_" & sName & "_MIFISH")
  END FUNCTION

  PUBLIC SUB ClearSessions()
    DIM oItem

    SESSION("FeedbackType_MIFISH")      = ""
    SESSION("Feedback_MIFISH")          = ""
    SESSION("FeedbackPage_MIFISH")      = ""
    SESSION("FeedbackHeading_MIFISH")   = ""
    SESSION("FeedbackAction_MIFISH")    = ""

    FOR EACH oItem IN SESSION.CONTENTS
      IF ( INSTR(oItem,"FeedbackFormField_") > 0 ) THEN
        Session.Contents.Remove(CSTR(oItem))
      END IF
    NEXT
  END SUB

  PUBLIC SUB setDefaults()
    c_Type            = 1
    c_Feedback        = ""
    c_FeedbackPage    = ""
    c_FeedbackHeading = ""
    c_FeedbackAction  = ""
  END SUB

  PRIVATE SUB Class_Initialize()
    CALL setDefaults()
  END SUB

  PRIVATE SUB Class_Terminate()
    CALL setDefaults()
  END SUB
END CLASS
%>
