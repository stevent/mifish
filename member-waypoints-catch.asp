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
' File:				    member-waypoints-catch.asp
' Version:			  1.0
' Author:			    Steven Taddei
' Modified By:		Steven Taddei
' Date Created:	  13/01/2016
' Last Modified:	13/01/2016
'
' Purpose:
'	Shows all catches for a specific
' user's waypoint
'***************************************
%>
<!--#include virtual="/admin/core/core.includes.asp"-->
<%
CALL Member.run("Login",NULL)
CALL Member.run("SecurePage",NULL)

'set SQL
DIM iPageSize       : iPageSize       = 10
DIM iCurrentPage    : iCurrentPage    = iReturnDefaultInt(REQUEST.QUERYSTRING("cp"),1)
DIM sOrderBy        : sOrderBy        = sReturnDefaultString(REQUEST.QUERYSTRING("sOrderBy"),"Title ASC")
DIM sQueryString    : sQueryString    = ""
DIM sDefaultURL     : sDefaultURL     = "member-waypoints-catch.asp"
DIM sCurrentURL     : sCurrentURL     = sDefaultURL
DIM sPaginationURL  : sPaginationURL  = sCurrentURL
DIM oWaypoint       : SET oWaypoint   = c_Waypoint.run("FindByID",REQUEST.QUERYSTRING("WaypointID"))
DIM oParams, oCatch, oFishes
DIM sFish

IF ( oWaypoint IS NOTHING ) THEN
  CALL Feedback.setError()
  CALL Feedback.setFeedbackHeading("No Waypoint Found")
  CALL Feedback.addToFeedback("<p>No waypoint was found to add the catch to.</p>")
  CALL Feedback.showFeedback("member-waypoints.asp")
END IF

'Set Class Variables
c_WaypointCatch.SetVar("PageSize")       = iPageSize
c_WaypointCatch.SetVar("CurrentPage")    = iCurrentPage

SET oParams     = oSetParams(ARRAY("conditions[]","order[]"))
oParams.ITEM("conditions")  = ARRAY("MemberID = " & setParamNumber(1) & " AND WaypointID = " & setParamNumber(2),Member.FieldValue("ID"),oWaypoint.FieldValue("ID"))

'create our find by params (if we need to)
IF ( bHaveInfo(sOrderBy) ) THEN
  oParams.ITEM("order") = ARRAY(sOrderBy)
END IF

'find out waypoints'
DIM oCatches      : SET oCatches    = c_WaypointCatch.run("FindAll",oParams)
DIM iRecordTotal  : iRecordTotal    = c_WaypointCatch.Var("RecordCount")
DIM iPageCount    : iPageCount      = c_WaypointCatch.Var("PageCount")

'Find our types
SET oFishes  = c_Fish.run("FindAll",NULL)

'set up query string
IF ( iReturnInt(iPageSize) > 0 ) THEN sQueryString = sAddToQueryString(sQueryString,"cp","[@CP@]")
IF ( bHaveInfo(sOrderBy) ) THEN sQueryString = sAddToQueryString(sQueryString,"sOrderBy",REQUEST.QUERYSTRING("sOrderBy"))

'set url variables'
sCurrentURL     = sCurrentURL & REPLACE(sQueryString,"[@CP@]",iCurrentPage)
sPaginationURL  = sPaginationURL & sQueryString
%>
<!DOCTYPE HTML>
<html lang="en">
  <head>
		<base href="<%= APPLICATION("SiteURL") %>" />
    <title>Catches | Wapoints | Members | MiFish Online</title>

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
					<h1>Catch Log : <%= oWaypoint.FieldValue("Title") %></h1>
				</header>
				<div class="box">
<%
IF ( Feedback.Has("Error") ) THEN
%>
          <div class="feedback error"><%= Feedback.Content %></div>
<%
ELSEIF ( Feedback.Has("Completed") ) THEN
%>
          <div class="feedback complete"><%= Feedback.Content %></div>
<%
END IF
%>
          <div class="table-wrapper">
            <h4><%= iRecordTotal %> fish caught</h4>

            <table class="alt">
              <thead>
                <tr>
                  <th>
                    <div class="order-by">
                      <a class="asc<% IF ( sOrderBy = "Title ASC" ) THEN RESPONSE.WRITE " selected" %>" href="<%= sDefaultURL %>?sOrderBy=Title ASC">ASC</a>
                      <a class="desc<% IF ( sOrderBy = "Title DESC" ) THEN RESPONSE.WRITE " selected" %>" href="<%= sDefaultURL %>?sOrderBy=Title DESC">DESC</a>
                    </div>
                    <span>Title</span>
                  </th>
                  <th>
                    <div class="order-by">
                      <a class="asc<% IF ( sOrderBy = "FishID ASC" ) THEN RESPONSE.WRITE " selected" %>" href="<%= sDefaultURL %>?sOrderBy=FishID ASC&WaypointID=<%= oWaypoint.FieldValue("ID") %>">ASC</a>
                      <a class="desc<% IF ( sOrderBy = "FishID DESC" ) THEN RESPONSE.WRITE " selected" %>" href="<%= sDefaultURL %>?sOrderBy=FishID DESC&WaypointID=<%= oWaypoint.FieldValue("ID") %>">DESC</a>
                    </div>
                    <span>Type</span>
                  </th>
                  <th>
                    <div class="order-by">
                      <a class="asc<% IF ( sOrderBy = "CatchDate ASC" ) THEN RESPONSE.WRITE " selected" %>" href="<%= sDefaultURL %>?sOrderBy=CatchDate ASC&WaypointID=<%= oWaypoint.FieldValue("ID") %>">ASC</a>
                      <a class="desc<% IF ( sOrderBy = "CatchDate Desc" ) THEN RESPONSE.WRITE " selected" %>" href="<%= sDefaultURL %>?sOrderBy=CatchDate DESC&WaypointID=<%= oWaypoint.FieldValue("ID") %>">DESC</a>
                    </div>
                    <span>Catch Date</span>
                  </th>
                  <th>
                    <div class="order-by">
                      <a class="asc<% IF ( sOrderBy = "Length ASC" ) THEN RESPONSE.WRITE " selected" %>" href="<%= sDefaultURL %>?sOrderBy=Length ASC&WaypointID=<%= oWaypoint.FieldValue("ID") %>">ASC</a>
                      <a class="desc<% IF ( sOrderBy = "Length DESC" ) THEN RESPONSE.WRITE " selected" %>" href="<%= sDefaultURL %>?sOrderBy=Length DESC&WaypointID=<%= oWaypoint.FieldValue("ID") %>">DESC</a>
                    </div>
                    <span>Length</span>
                  </th>
                  <th>
                    <div class="order-by">
                      <a class="asc<% IF ( sOrderBy = "Weight ASC" ) THEN RESPONSE.WRITE " selected" %>" href="<%= sDefaultURL %>?sOrderBy=Weight ASC&WaypointID=<%= oWaypoint.FieldValue("ID") %>">ASC</a>
                      <a class="desc<% IF ( sOrderBy = "Weight DESC" ) THEN RESPONSE.WRITE " selected" %>" href="<%= sDefaultURL %>?sOrderBy=Weight DESC&WaypointID=<%= oWaypoint.FieldValue("ID") %>">DESC</a>
                    </div>
                    <span>Weight</span>
                  </th>
                  <th>&nbsp;</th>
                </tr>
              </thead>
              <tbody>
<%
FOR EACH oCatch IN oCatches.ITEMS
  'default
  sFish = ""

  IF ( oFishes.EXISTS(CSTR(oCatch.FieldValue("FishID"))) ) THEN
    sFish = oFishes.ITEM(CSTR(oCatch.FieldValue("FishID"))).FieldValue("Name")
  END IF
%>
                <tr>
                  <td><%= oCatch.FieldValue("Title") %></td>
                  <td><%= sFish %></td>
                  <td><%= oCatch.FieldValue("CatchDate") %></td>
                  <td><%= oCatch.FieldValue("Length") %></td>
                  <td><%= oCatch.FieldValue("Weight") %></td>
                  <td><a href="#" title="View Waypoint">view</a> | <a href="member-waypoints-catch-form.asp?id=<%= oCatch.FieldValue("ID") %>&WaypointID=<%= oWaypoint.FieldValue("ID") %>" title="Edit Catch">edit</a> | <a href="member-waypoints-catch-action.asp?id=<%= oCatch.FieldValue("ID") %>&WaypointID=<%= oWaypoint.FieldValue("ID") %>&action=del" title="Delete Catch">delete</a></td>
                </tr>
<%
NEXT
%>
              </tbody>
              <tfoot>
                <tr>
                  <td colspan="6">
<%
IF ( oCatches.COUNT = 0 ) THEN
%>
                    <div class="feedback">No fish has been logged for this location</div>
<%
END IF
%>
                    <% CALL DisplayPagingNav(sPaginationURL,iPageSize,iCurrentPage,iRecordTotal) %>

                    <div class="actions">
                      <a class="button small" href="member-waypoints-catch-form.asp?WaypointID=<%= oWaypoint.FieldValue("ID") %>">add catch</a>
                    </div>
                  </td>
                </tr>
              </tfoot>
            </table>
          </div>
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
