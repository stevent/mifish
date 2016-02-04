<%
OPTION EXPLICIT

'Ensure a new copy of the page every time
RESPONSE.EXPIRES	= -1
RESPONSE.BUFFER	= TRUE

'***************************************
' Copyright (c) 2001-2008 Webfirm
' http://www.webfirm.com.au
' Ph: 1300 304 779
'
' File:				    member-waypoints.asp
' Version:			  1.0
' Author:			    Steven Taddei
' Modified By:		Steven Taddei
' Date Created:	  13/01/2016
' Last Modified:	13/01/2016
'
' Purpose:
'	Shows all waypoints for the specific
' user
'***************************************
%>
<!--#include virtual="/admin/core/core.includes.asp"-->
<%
CALL Member.run("Login",NULL)
CALL Member.run("SecurePage",NULL)

DIM oTypes        : SET oTypes      = c_WaypointType.run("FindAll",NULL)
DIM oWaypoint     : SET oWaypoint   = c_Waypoint.run("FindByID",REQUEST.QUERYSTRING("id"))
DIM oSounders     : SET oSounders   = c_Sounder.run("FindByMember",Member.FieldValue("ID"))

DIM sPageTitle, sHeading, sSelected, sChecked, sFieldLabel, sFieldName, sFieldValue, sFieldID, sType, sCoordLong, sCoordLat
DIM sDegrees, sMinutes, sDir
DIM oType, oSounder, oSounderFields, oSounderValues, oField, oCoOrd

sPageTitle  = "Waypoint Not Found | Waypoints | Members | MiFish Online"
sHeading    = "Waypoint Not Found"

IF ( NOT oWaypoint IS NOTHING ) THEN
  sPageTitle  = oWaypoint.FieldValue("Title") & " | Waypoints | Members | MiFish Online"
  sHeading    = oWaypoint.FieldValue("Title")
END IF
%>
<!DOCTYPE HTML>
<html lang="en">
  <head>
		<base href="<%= APPLICATION("SiteURL") %>" />
    <title><%= sPageTitle %></title>

		<!--#include virtual="/assets/views/scripts-styles.asp"-->
	</head>
	<body class="popup">
		<div id="page-wrapper">
			<!-- Main -->
			<section id="main" class="container 100%">
				<header>
					<h1><%= sHeading %></h1>
				</header>
				<div class="box">
<%
IF ( NOT oWaypoint IS NOTHING ) THEN
  'default
  sType = ""

  IF ( oTypes.EXISTS(CSTR(oWaypoint.FieldValue("Type"))) ) THEN
    sType = oTypes.ITEM(CSTR(oWaypoint.FieldValue("Type"))).FieldValue("Name")
  END IF

  sDegrees    = ""
  sMinutes    = ""
  sDir        = "E"
  SET oCoOrd  = oConvertToDegree(oWaypoint.FieldValue("Longitude"),"lon")

  IF ( bHaveInfo(oCoOrd.ITEM("degrees")) ) THEN sDegrees = oCoOrd.ITEM("degrees")
  IF ( bHaveInfo(oCoOrd.ITEM("minutes")) ) THEN sMinutes = FORMATNUMBER(oCoOrd.ITEM("minutes"),3)
  IF ( bHaveInfo(oCoOrd.ITEM("dir")) ) THEN sDir = oCoOrd.ITEM("dir")

  IF ( bHaveInfo(sDegrees) AND bHaveInfo(sMinutes) AND bHaveInfo(sDir) ) THEN
    sCoordLong = sDegrees & "°" & sMinutes & "'" & sDir
  END IF

  sDegrees    = ""
  sMinutes    = ""
  sDir        = "S"
  SET oCoOrd  = oConvertToDegree(oWaypoint.FieldValue("Latitude"),"lat")

  IF ( bHaveInfo(oCoOrd.ITEM("degrees")) ) THEN sDegrees = oCoOrd.ITEM("degrees")
  IF ( bHaveInfo(oCoOrd.ITEM("minutes")) ) THEN sMinutes = FORMATNUMBER(oCoOrd.ITEM("minutes"),3)
  IF ( bHaveInfo(oCoOrd.ITEM("dir")) ) THEN sDir = oCoOrd.ITEM("dir")

  IF ( bHaveInfo(sDegrees) AND bHaveInfo(sMinutes) AND bHaveInfo(sDir) ) THEN
    sCoordLat = sDegrees & "°" & sMinutes & "'" & sDir
  END IF
%>
          <div class="table-wrapper">
            <table class="single-data">
              <tbody>
                <tr>
                  <td>Type</td>
                  <td><%= sType %></td>
                </tr>
                <tr>
                  <td>Longitude</td>
                  <td><%= oWaypoint.FieldValue("Longitude") %></td>
                </tr>
                <tr>
                  <td>Latitude</td>
                  <td><%= oWaypoint.FieldValue("Latitude") %></td>
                </tr>
                <tr>
                  <td>Co-Ordinates</td>
                  <td>
<%
  IF ( bHaveInfo(sCoordLong) AND bHaveInfo(sCoordLat) ) THEN
%>
                    <span class="lon"><%= sCoordLong %></span>&nbsp;<span class="lon"><%= sCoordLat %></span>
<%
  END IF
%>
                  </td>
                </tr>
                <tr>
                  <td>Notes</td>
                  <td><%= oWaypoint.FieldValue("Notes") %></td>
                </tr>
              </tbody>
            </table>
          </div>
<%
ELSE
%>
          <div class="feedback error">The selected waypoint could not be found</div>
<%
END IF
%>
				</div>
			</section>
		</div>

		<!--#include virtual="/assets/views/footer-scripts.asp"-->
	</body>
</html>
