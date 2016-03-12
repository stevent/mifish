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
' File:				    member-waypoints-form.asp
' Version:			  1.0
' Author:			    Steven Taddei
' Modified By:		Steven Taddei
' Date Created:	  1515/01/2016
' Last Modified:	12/01/2016
'
' Purpose:
'	Allows user to edit or add a waypoint
'***************************************
%>
<!--#include virtual="/admin/core/core.includes.asp"-->
<%
CALL Member.run("Login",NULL)
CALL Member.run("SecurePage",NULL)

DIM iPageSize     : iPageSize       = 10
DIM oTypes        : SET oTypes      = c_WaypointType.run("FindAll",NULL)
DIM oWaypoint     : SET oWaypoint   = c_Waypoint.run("FindByID",REQUEST.QUERYSTRING("id"))
DIM oSounders     : SET oSounders   = c_Sounder.run("FindByMember",Member.FieldValue("ID"))
DIM sAction       : sAction         = "update"
DIM sPageTitle, sHeading, sSelected, sChecked, sFieldLabel, sFieldName, sFieldValue, sFieldID, sBtnText, sBtnTextAlt
DIM sDegrees, sMinutes, sDir
DIM oType, oSounder, oSounderFields, oSounderValues, oField, oCoOrd

IF ( oWaypoint IS NOTHING ) THEN
  SET oWaypoint   = c_Waypoint.run("New",NULL)
  sAction         = "create"

  'defaults
  oWaypoint.SetValue("Type") = 1
END IF

SELECT CASE UCASE(sAction)
  CASE "CREATE"
    sPageTitle  = "Create Waypoint | Members | MiFish Online"
    sHeading    = "Create Waypoint"
    sBtnText    = "Create"
    sBtnTextAlt = "Create and add another"

  CASE "UPDATE"
    sPageTitle  = "Update Waypoint | Members | MiFish Online"
    sHeading    = "Update Waypoint"
    sBtnText    = "Update"
    sBtnTextAlt = "Update and continue edit"

END SELECT
%>
<!DOCTYPE HTML>
<html lang="en">
  <head>
		<base href="<%= APPLICATION("SiteURL") %>" />
    <title><%= sPageTitle %></title>

		<!--#include virtual="/assets/views/scripts-styles.asp"-->
	</head>
	<body>
		<div id="page-wrapper">
			<!-- Header -->
			<header id="header">
				<!--#include virtual="/assets/views/header.asp"-->
			</header>

			<!-- Main -->
			<section id="main" class="container 100%">
				<header>
					<h1><%= sHeading %></h1>
				</header>
				<div class="box">
<%
IF ( Feedback.Has("Error") ) THEN
  oWaypoint.SetValue("Title") = Feedback.getField("Title")
  oWaypoint.SetValue("Type") = Feedback.getField("Type")
  oWaypoint.SetValue("Notes") = Feedback.getField("Notes")

  SET oCoOrd  = oConvertToDegree(NULL,"lon")
  oCoOrd.ITEM("degrees")          = Feedback.getField("LongitudeDegrese")
  oCoOrd.ITEM("minutes")          = Feedback.getField("LongitudeMinutes")
  oCoOrd.ITEM("dir")              = Feedback.getField("LongitudeDir")
  oWaypoint.SetValue("Longitude") = iConvertToDecimal(oCoOrd,"lon")

  SET oCoOrd  = oConvertToDegree(NULL,"lat")
  oCoOrd.ITEM("degrees")          = Feedback.getField("LatitudeDegrees")
  oCoOrd.ITEM("minutes")          = Feedback.getField("LatitudeMinutes")
  oCoOrd.ITEM("dir")              = Feedback.getField("LatitudeDir")
  oWaypoint.SetValue("Latitude")  = iConvertToDecimal(oCoOrd,"lat")
%>
          <div class="feedback error"><%= Feedback.Content %></div>
<%
ELSEIF ( Feedback.Has("Completed") ) THEN
%>
          <div class="feedback complete"><%= Feedback.Content %></div>
<%
END IF
%>
					<form id="waypoints" action="<%= APPLICATION("SiteURL") %>/member-waypoints-action.asp" method="post">
            <div class="hidden_fields">
              	<input type="text" name="Id" id="waypoints_id" value="<%= oWaypoint.FieldValue("ID") %>" />
                <input type="text" name="Action" id="waypoints_action" value="<%= sAction %>" />
            </div>
						<div class="row uniform 50%">
              <!-- Type, Title, Longitude, Latitude, Notes -->
							<div class="12u">
                <label>Title</label>
								<input type="text" name="Title" id="waypoints_title" placeholder="Title" value="<%= oWaypoint.FieldValue("Title") %>" />
							</div>
						</div>
						<div class="row uniform 50%">
							<div class="12u">
                <label>Type</label>
                <select name="Type" id="waypoints_type">
                  <option value="">Select Type</option>
<%
FOR EACH oType IN oTypes.ITEMS
  'default
  sSelected = ""
  IF ( sReturnString(oType.FieldValue("ID")) = sReturnString(oWaypoint.FieldValue("Type")) ) THEN sSelected = " selected=""selected"""
%>
                  <option value="<%= oType.FieldValue("ID") %>"<%= sSelected %>><%= oType.FieldValue("Name") %></option>
<%
NEXT
%>
                </select>
							</div>
						</div>
<%
sDegrees    = ""
sMinutes    = ""
sDir        = "E"
SET oCoOrd  = oConvertToDegree(oWaypoint.FieldValue("Longitude"),"lon")

IF ( bHaveInfo(oCoOrd.ITEM("degrees")) ) THEN sDegrees = oCoOrd.ITEM("degrees")
IF ( bHaveInfo(oCoOrd.ITEM("minutes")) ) THEN sMinutes = FORMATNUMBER(oCoOrd.ITEM("minutes"),3)
IF ( bHaveInfo(oCoOrd.ITEM("dir")) ) THEN sDir = oCoOrd.ITEM("dir")
%>
            <div class="row uniform 50%">
              <div class="12u">
                <label>Longitude</label>
              </div>
            </div>
            <div class="row uniform 50%">
							<div class="3u">
                <input type="text" name="LongitudeDegrese" id="waypoints_longitude_degrees" placeholder="XXX" value="<%= sDegrees %>" />
							</div>
              <div class="1u"><h4>°</h4></div>
              <div class="3u">
                <input type="text" name="LongitudeMinutes" id="waypoints_longitude_minutes" placeholder="XX.XXX" value="<%= sMinutes %>" />
							</div>
              <div class="1u">'</div>
              <div class="3u">
                <select name="LongitudeDir" id="waypoints_longitude_dir">
                  <option value="N"<% IF ( UCASE(sDir) = "E" ) THEN RESPONSE.WRITE " selected=""selected""" %>>E</option>
                  <option value="S"<% IF ( UCASE(sDir) = "W" ) THEN RESPONSE.WRITE " selected=""selected""" %>>W</option>
                </select>
							</div>
						</div>
<%
sDegrees    = ""
sMinutes    = ""
sDir        = "S"
SET oCoOrd  = oConvertToDegree(oWaypoint.FieldValue("Latitude"),"lat")

IF ( bHaveInfo(oCoOrd.ITEM("degrees")) ) THEN sDegrees = oCoOrd.ITEM("degrees")
IF ( bHaveInfo(oCoOrd.ITEM("minutes")) ) THEN sMinutes = FORMATNUMBER(oCoOrd.ITEM("minutes"),3)
IF ( bHaveInfo(oCoOrd.ITEM("dir")) ) THEN sDir = oCoOrd.ITEM("dir")
%>
            <div class="row uniform 50%">
              <div class="12u">
                <label>Latitude</label>
              </div>
            </div>
            <div class="row uniform 50%">
							<div class="3u">
                <input type="text" name="LatitudeDegrees" id="waypoints_latitude_degrees" placeholder="XX" value="<%= sDegrees %>" />
							</div>
              <div class="1u"><h4>°</h4></div>
              <div class="3u">
                <input type="text" name="LatitudeMinutes" id="waypoints_latitude_minutes" placeholder="XX.XXX" value="<%= sMinutes %>" />
							</div>
              <div class="1u">'</div>
              <div class="3u">
                <select name="LatitudeDir" id="waypoints_latitude_dir">
                  <option value="N"<% IF ( UCASE(sDir) = "N" ) THEN RESPONSE.WRITE " selected=""selected""" %>>N</option>
                  <option value="S"<% IF ( UCASE(sDir) = "S" ) THEN RESPONSE.WRITE " selected=""selected""" %>>S</option>
                </select>
							</div>
						</div>
            <div class="row uniform 50%">
							<div class="12u">
                <label>Notes</label>
                  <textarea name="Notes" rews="10" cols="100" id="waypoints_notes"><%= oWaypoint.FieldValue("Notes") %></textarea>
							</div>
						</div>

            <hr />

            <div class="row uniform">
							<div class="12u">
								<ul class="actions align-center">
									<li><input type="submit" value="<%= sBtnText %>" class="disable-on-click change-action" data-action="<%= sAction %>" /></li>
                  <li><input type="submit" value="<%= sBtnTextAlt %>" class="disable-on-click change-action" data-action="<%= sAction %>_form"  /></li>
								</ul>
							</div>
            </div>

            <hr />
<%
FOR EACH oSounder IN oSounders.ITEMS
%>
            <h3><%= oSounder.FieldValue("Name") %></h3>
<%

  SET oSounderFields = c_SounderField.run("FindBySounder",oSetParams(ARRAY("iSounderID[" & oSounder.FieldValue("ID") & "]","iWaypointID[" & oWaypoint.FieldValue("ID") & "]")))

  FOR EACH oField IN oSounderFields.ITEMS
    'only show if its not linked to a waypoint field
    IF ( NOT bHaveInfo(oField.FieldValue("WaypointField")) ) THEN
      sFieldName    = "SounderField_" & oField.FieldValue("ID")
      sFieldLabel   = oField.FieldValue("Name")
      sFieldValue   = oField.FieldValue("FieldValue")
      sFieldID      = "SounderField_" & oField.FieldValue("ID")
%>
            <div class="row uniform 50%">
              <div class="12u">
                <label><%= sFieldLabel %></label>
                <input type="text" name="<%= sFieldName %>" id="<%= sFieldID %>" placeholder="<%= sFieldLabel %>" value="<%= sFieldValue %>" />
              </div>
            </div>
<%
    END IF
  NEXT
%>
            <hr />
<%
NEXT
%>
						<div class="row uniform">
							<div class="12u">
								<ul class="actions align-center">
									<li><input type="submit" value="<%= sBtnText %>" class="disable-on-click change-action" data-action="<%= sAction %>" /></li>
                  <li><input type="submit" value="<%= sBtnTextAlt %>" class="disable-on-click change-action" data-action="<%= sAction %>_form"  /></li>
								</ul>
							</div>
						</div>
					</form>
				</div>
			</section>

			<!-- Footer -->
			<footer id="footer">
				<!--#include virtual="/assets/views/footer.asp"-->
			</footer>
		</div>

		<!--#include virtual="/assets/views/footer-scripts.asp"-->
	</body>
</html>
