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

'set SQL
DIM iPageSize       : iPageSize       = 10
DIM iCurrentPage    : iCurrentPage    = iReturnDefaultInt(REQUEST.QUERYSTRING("cp"),1)
DIM sOrderBy        : sOrderBy        = sReturnDefaultString(REQUEST.QUERYSTRING("sOrderBy"),"Title ASC")
DIM sQueryString    : sQueryString    = ""
DIM sDefaultURL     : sDefaultURL     = "member-waypoints.asp"
DIM sCurrentURL     : sCurrentURL     = sDefaultURL
DIM sPaginationURL  : sPaginationURL  = sCurrentURL
DIM oParams, oWaypoint, oTypes, oCatches
DIM sType

'Set Class Variables
c_Waypoint.SetVar("PageSize")       = iPageSize
c_Waypoint.SetVar("CurrentPage")    = iCurrentPage

'create our find by params (if we need to)
IF ( bHaveInfo(sOrderBy) ) THEN
  SET oParams     = oSetParams(ARRAY("conditions[]","order[]"))

  oParams.ITEM("conditions")  = ARRAY("MemberID = " & setParamNumber(1),Member.FieldValue("ID"))
  oParams.ITEM("order")       = ARRAY(sOrderBy)
ELSE
  oParams = Member.FieldValue("ID")
END IF

'find out waypoints'
DIM oWaypoints    : SET oWaypoints  = c_Waypoint.run("FindByMember",oParams)
DIM iRecordTotal  : iRecordTotal    = c_Waypoint.Var("RecordCount")
DIM iPageCount    : iPageCount      = c_Waypoint.Var("PageCount")

'Find our types
SET oTypes  = c_WaypointType.run("FindAll",NULL)

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
    <title>Wapoints | Members | MiFish Online</title>

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
					<h1>Waypoints</h1>
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
            <h4><%= iRecordTotal %> waypoints found</h4>

            <table class="alt">
              <thead>
                <tr>
                  <th>
                    <div class="order-by">
                      <a class="asc<% IF ( sOrderBy = "Title ASC" ) THEN RESPONSE.WRITE " selected" %>" href="<%= sDefaultURL %>?sOrderBy=Title ASC">ASC</a>
                      <a class="desc<% IF ( sOrderBy = "Title DESC" ) THEN RESPONSE.WRITE " selected" %>" href="<%= sDefaultURL %>?sOrderBy=Title DESC">DESC</a>
                    </div>
                    <span>Name</span>
                  </th>
                  <th>
                    <div class="order-by">
                      <a class="asc<% IF ( sOrderBy = "Type ASC" ) THEN RESPONSE.WRITE " selected" %>" href="<%= sDefaultURL %>?sOrderBy=Type ASC">ASC</a>
                      <a class="desc<% IF ( sOrderBy = "Type DESC" ) THEN RESPONSE.WRITE " selected" %>" href="<%= sDefaultURL %>?sOrderBy=Type DESC">DESC</a>
                    </div>
                    <span>Type</span>
                  </th>
                  <th>
                    <div class="order-by">
                      <a class="asc<% IF ( sOrderBy = "Longitude ASC" ) THEN RESPONSE.WRITE " selected" %>" href="<%= sDefaultURL %>?sOrderBy=Longitude ASC">ASC</a>
                      <a class="desc<% IF ( sOrderBy = "Longitude Desc" ) THEN RESPONSE.WRITE " selected" %>" href="<%= sDefaultURL %>?sOrderBy=Longitude DESC">DESC</a>
                    </div>
                    <span>Longitude</span>
                  </th>
                  <th>
                    <div class="order-by">
                      <a class="asc<% IF ( sOrderBy = "Latitude ASC" ) THEN RESPONSE.WRITE " selected" %>" href="<%= sDefaultURL %>?sOrderBy=Latitude ASC">ASC</a>
                      <a class="desc<% IF ( sOrderBy = "Latitude DESC" ) THEN RESPONSE.WRITE " selected" %>" href="<%= sDefaultURL %>?sOrderBy=Latitude DESC">DESC</a>
                    </div>
                    <span>Latitude</span>
                  </th>
                  <th>
                    <span>Notes</span>
                  </th>
                  <th>&nbsp;</th>
                </tr>
              </thead>
              <tbody>
<%
FOR EACH oWaypoint IN oWaypoints.ITEMS
  'default
  sType = ""

  IF ( oTypes.EXISTS(CSTR(oWaypoint.FieldValue("Type"))) ) THEN
    sType = oTypes.ITEM(CSTR(oWaypoint.FieldValue("Type"))).FieldValue("Name")
  END IF

  'set catches params
  SET oParams     = oSetParams(ARRAY("conditions[]","order[]"))
  oParams.ITEM("conditions")  = ARRAY("MemberID = " & setParamNumber(1) & " AND WaypointID = " & setParamNumber(2),Member.FieldValue("ID"),oWaypoint.FieldValue("ID"))

  'find out waypoints'
  SET oCatches = c_WaypointCatch.run("FindAll",oParams)
%>
                <tr>
                  <td><%= oWaypoint.FieldValue("Title") %></td>
                  <td><%= sType %></td>
                  <td><%= oWaypoint.FieldValue("Longitude") %></td>
                  <td><%= oWaypoint.FieldValue("Latitude") %></td>
                  <td><%= oWaypoint.FieldValue("Notes") %></td>
                  <td><a href="member-waypoints-catch.asp?WaypointID=<%= oWaypoint.FieldValue("ID") %>" title="Fish Caught (0)">fish caught (<%= oCatches.COUNT %>)</a> | <a href="member-waypoints-catch-form.asp?WaypointID=<%= oWaypoint.FieldValue("ID") %>" title="Add Catch">add catch</a> | <a href="#" title="View Waypoint">view</a> | <a href="member-waypoints-form.asp?id=<%= oWaypoint.FieldValue("ID") %>" title="Edit Waypoint">edit</a> | <a href="member-waypoints-action.asp?id=<%= oWaypoint.FieldValue("ID") %>&action=del" title="Deete Waypoint">delete</a></td>
                </tr>
<%
NEXT
%>
              </tbody>
              <tfoot>
                <tr>
                  <td colspan="6">
<%
IF ( oWaypoints.COUNT = 0 ) THEN
%>
                    <div class="feedback">You currently have no waypoints</div>
<%
END IF
%>
                    <% CALL DisplayPagingNav(sPaginationURL,iPageSize,iCurrentPage,iRecordTotal) %>

                    <div class="actions">
                      <a class="button small" href="member-waypoints-import.asp">import waypoint</a>
                      <a class="button small" href="member-waypoints-form.asp">add waypoint</a>
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
