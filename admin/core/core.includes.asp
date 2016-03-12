<!--#include file="adovbs.asp"-->
<!--#include file="helpers.application.asp"-->
<!--#include file="class.error.asp"-->
<!--#include file="core.request.asp"-->
<!--#include file="core.feedback.asp"-->
<!--#include file="class.field.asp"-->
<!--#include file="db.connection.asp"-->
<!--#include file="db.recordset.asp"-->
<!--#include file="db.like.recordset.asp"-->
<!--#include file="class.klass.asp"-->
<!--#include file="class.base.asp"-->
<!--#include file="db.module.asp"-->

<!--#include file="plugins/sha256/sha256.asp"-->

<!--#include file="../module/helpers.asp"-->
<!--#include file="../module/page/class.asp"-->
<!--#include file="../module/member/class.asp"-->
<!--#include file="../module/waypoint/class.asp"-->
<!--#include file="../module/waypoint_type/class.asp"-->
<!--#include file="../module/sounder/class.asp"-->
<!--#include file="../module/sounder_field/class.asp"-->
<!--#include file="../module/fish/class.asp"-->
<!--#include file="../module/waypoint_catch/class.asp"-->

<!--#include file="../module/waypoint/helpers.asp"-->
<!--#include file="../module/waypoint_catch/helpers.asp"-->
<%
'initialize global variables
DIM p_oFso, p_oFolder, p_oSubFolder

'
'LETS DYNAMICALLY INCLUDE ALL MODULE CLASSES
'
SET p_oFso    = SERVER.CREATEOBJECT("Scripting.FileSystemObject")
SET p_oFolder = p_oFso.GETFOLDER(SERVER.MAPPATH("\admin\module"))

FOR EACH p_oSubFolder IN p_oFolder.SUBFOLDERS
  IF ( file_exists("\admin\module\" & p_oSubFolder.Name & "\class.asp") ) THEN
    'EXECUTE(grabInclude("\admin\module\" & p_oSubFolder.Name & "\class.asp"))
  END IF
NEXT

SET p_oFso    = NOTHING
SET p_oFolder = NOTHING
%>
