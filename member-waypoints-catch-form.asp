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
' File:				    member-waypoints-catch-form.asp
' Version:			  1.0
' Author:			    Steven Taddei
' Modified By:		Steven Taddei
' Date Created:	  1515/01/2016
' Last Modified:	12/01/2016
'
' Purpose:
'	Allows user to edit or add a catch on a
' specific waypoint
'***************************************
%>
<!--#include virtual="/admin/core/core.includes.asp"-->
<%
CALL Member.run("Login",NULL)
CALL Member.run("SecurePage",NULL)

DIM oTypes        : SET oTypes      = c_WaypointType.run("FindAll",NULL)
DIM oCatch        : SET oCatch      = c_WaypointCatch.run("FindByID",REQUEST.QUERYSTRING("id"))
DIM oWaypoint     : SET oWaypoint   = c_Waypoint.run("FindByID",REQUEST.QUERYSTRING("WaypointID"))
DIM oFishes       : SET oFishes     = c_Fish.run("FindAll",NULL)
DIM sAction       : sAction         = "update"
DIM sPageTitle, sHeading, sBtnText, sBtnTextAlt, sSelected
DIM oFish

IF ( oWaypoint IS NOTHING ) THEN
  CALL Feedback.setError()
  CALL Feedback.setFeedbackHeading("No Waypoint Found")
  CALL Feedback.addToFeedback("<p>No waypoint was found to add the catch to.</p>")
  CALL Feedback.showFeedback("member-waypoints.asp")
END IF

IF ( oCatch IS NOTHING ) THEN
  SET oCatch      = c_WaypointCatch.run("New",NULL)
  sAction         = "create"
ELSE
  IF ( INT(oWaypoint.FieldValue("ID")) <> INT(oCatch.FieldValue("WaypointID")) ) THEN
    RESPONSE.REDIRECT APPLICATION("SiteURL") & "/member-waypoints-catch-form.asp?id=" & oCatch.FieldValue("ID") & "&WaypointID=" & oCatch.FieldValue("WaypointID")
  END IF
END IF

SELECT CASE UCASE(sAction)
  CASE "CREATE"
    sPageTitle  = "Add Catch Log | Members | MiFish Online"
    sHeading    = "Add Catch Log : " &  oWaypoint.FieldValue("Title")
    sBtnText    = "Create"
    sBtnTextAlt = "Create and add another"

  CASE "UPDATE"
    sPageTitle  = "Edit Catch Log | Members | MiFish Online"
    sHeading    = "Edit Catch Log : " &  oWaypoint.FieldValue("Title")
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
  oCatch.SetValue("Title") = Feedback.getField("Title")
  oCatch.SetValue("Type")  = Feedback.getField("Type")
  oCatch.SetValue("Notes") = Feedback.getField("Notes")
%>
          <div class="feedback error"><%= Feedback.Content %></div>
<%
ELSEIF ( Feedback.Has("Completed") ) THEN
%>
          <div class="feedback complete"><%= Feedback.Content %></div>
<%
END IF
%>
					<form id="waypoints" action="<%= APPLICATION("SiteURL") %>/member-waypoints-catch-action.asp" method="post">
            <div class="hidden_fields">
                <input type="text" name="Action" id="waypoints_action" value="<%= sAction %>" />
            </div>
            <div class="row uniform 50%">
							<div class="12u">
                <label>Waypoint</label>
								<input type="text" name="WaypointTitle" id="waypoint_catch_waypointtitle" placeholder="" value="<%= oWaypoint.FieldValue("Title") %>" class="readonly" readonly="readonly" />
                <input type="hidden" name="WaypointID" id="waypoint_catch_waypointid" placeholder="" value="<%= oWaypoint.FieldValue("ID") %>" />
							</div>
						</div>
						<div class="row uniform 50%">
              <!-- Type, Title, Longitude, Latitude, Notes -->
							<div class="12u">
                <label>Title</label>
								<input type="text" name="Title" id="waypoint_catch_title" placeholder="Title" value="<%= oCatch.FieldValue("Title") %>" />
							</div>
						</div>
						<div class="row uniform 50%">
							<div class="12u">
                <label>Fish Type</label>
                <select name="FishID" id="waypoint_catch_fishid">
                  <option value="">Select Fish Type</option>
<%
FOR EACH oFish IN oFishes.ITEMS
  'default
  sSelected = ""
  IF ( sReturnString(oFish.FieldValue("ID")) = sReturnString(oCatch.FieldValue("FishID")) ) THEN sSelected = " selected=""selected"""
%>
                  <option value="<%= oFish.FieldValue("ID") %>"<%= sSelected %>><%= oFish.FieldValue("Name") %></option>
<%
NEXT
%>
                </select>
							</div>
						</div>
            <div class="row uniform 50%">
              <!-- Type, Title, Longitude, Latitude, Notes -->
							<div class="12u">
                <label>Catch Date</label>
								<input type="text" name="CatchDate" id="waypoint_catch_catch_date" placeholder="YYYY-MM-DD" value="<%= dGetFormattedDate( "Y-m-d", oCatch.FieldValue("CatchDate") ) %>" />
							</div>
						</div>
            <div class="row uniform 50%">
							<div class="12u">
                <label>Location Info</label>
								<input type="text" name="Location" id="waypoint_catch_location" placeholder="Location Info" value="<%= oCatch.FieldValue("Location") %>" />
							</div>
						</div>
            <div class="row uniform 50%">
							<div class="12u">
                <label>Time Of Day Info</label>
								<input type="text" name="TimeOfDay" id="waypoint_catch_timeofday" placeholder="Time of Day" value="<%= oCatch.FieldValue("TimeOfDay") %>" />
							</div>
						</div>
            <div class="row uniform 50%">
							<div class="12u">
                <label>Length</label>
								<input type="text" name="Length" id="waypoint_catch_length" placeholder="Length" value="<%= oCatch.FieldValue("Length") %>" />
							</div>
						</div>
            <div class="row uniform 50%">
							<div class="12u">
                <label>Weight</label>
								<input type="text" name="Weight" id="waypoint_catch_weight" placeholder="Weight" value="<%= oCatch.FieldValue("Weight") %>" />
							</div>
						</div>
            <div class="row uniform 50%">
							<div class="12u">
                <label>Notes</label>
                  <textarea name="Notes" rews="10" cols="100" id="waypoint_catch_notes"><%= oCatch.FieldValue("Notes") %></textarea>
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
