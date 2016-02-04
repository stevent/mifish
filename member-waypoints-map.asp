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
DIM iPageSize     : iPageSize       = 10
DIM oTypes        : SET oTypes      = c_WaypointType.run("FindAll",NULL)
DIM oParams
DIM sOrderBy

sOrderBy = sReturnDefaultString(REQUEST.QUERYSTRING("sOrderBy"),"Title ASC")

IF ( bHaveInfo(sOrderBy) ) THEN
  SET oParams     = oSetParams(ARRAY("conditions[]","order[]"))

  oParams.ITEM("conditions")  = ARRAY("MemberID = " & setParamNumber(1),Member.FieldValue("ID"))
  oParams.ITEM("order")       = ARRAY(sOrderBy)
ELSE
  oParams = Member.FieldValue("ID")
END IF

DIM oWaypoints    : SET oWaypoints  = c_Waypoint.run("FindByMember",oParams)
DIM iRecordTotal  : iRecordTotal    = oWaypoints.COUNT
DIM oWaypoint
DIM sType
%>
<!DOCTYPE HTML>
<html lang="en">
  <head>
		<base href="<%= APPLICATION("SiteURL") %>" />
    <title>Wapoints Map View | Members | MiFish Online</title>

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
					<h1>Map</h1>
				</header>
				<div class="box">
          <div id="map-toolbar">
            <a href="#" class="hide_all">hide all</a> <a href="#" class="show_all">show all</a>
          </div>
          <div class="marker_container clearfix" style="width:100%; height: 400px; z-index: 100; position: relative;">
            <div id="marker-list" style="width: 30%; float: left; height: 400px; background: #FFF; border: 1px solid #777; border-right: none; overflow: auto; overflow-x: hidden;">
                <ul class="alt" style="padding: 2.5px;">

                </ul>
            </div>
            <div id="map" style="width:70%; height: 400px; position: relative; float: left; border: 1px solid #777;"></div>
          </div>
				</div>
			</section>

			<!-- Footer -->
			<footer id="footer">
				<!--#include virtual="/assets/views/footer.asp"-->
			</footer>
		</div>

		<!--#include virtual="/assets/views/footer-scripts.asp"-->
    <script type="text/javascript" src="https://maps.googleapis.com/maps/api/js"></script>
    <script type="text/javascript" src="/assets/js/marker.cluster-comp.js"></script>
    <script type="text/javascript" src="/cache/<%= MemberWptJsonFilename %>"></script>
    <script type="text/javascript" src="/assets/js/waypoint.marker.map.js"></script>
	</body>
</html>
