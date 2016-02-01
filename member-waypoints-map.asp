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
					<h1>Waypoints</h1>
				</header>
				<div class="box">
          <p>map test</p>
          <div id="map" style="width:100%; height: 400px;"></div>
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
    <script type="text/javascript" src="/cache/<%= JsonWptFilename %>"></script>

    <script type="text/javascript">

      var markerClusterer = null;
      var map = null;
      var imageUrl = 'http://chart.apis.google.com/chart?cht=mm&chs=24x32&' +
          'chco=FFFFFF,008CFF,000000&ext=.png';

      function refreshMap() {
        if (markerClusterer) {
          markerClusterer.clearMarkers();
        }

        var markers = [];

        var markerImage = new google.maps.MarkerImage(imageUrl,
          new google.maps.Size(24, 32));

        if ( data.items.length > 0 ) {
          var latlngbounds = new google.maps.LatLngBounds();

          for (var i = 0; i < data.items.length; ++i) {

            var latLng = new google.maps.LatLng(data.items[i].lat,data.items[i].lon)

            latlngbounds.extend(latLng);

            var marker = new google.maps.Marker({
              position: latLng,
              draggable: false,
              icon: markerImage
            });
            markers.push(marker);
          }

          map.fitBounds(latlngbounds);
        }

        var size = 5

        if ( markers.length > 0 ) {
          markerClusterer = new MarkerClusterer(map, markers, {
            gridSize: size
          });
        }
      }

      function initialize() {
        map = new google.maps.Map(document.getElementById('map'), {
          zoom: 2,
          center: new google.maps.LatLng(39.91, 116.38),
          mapTypeId: google.maps.MapTypeId.ROADMAP
        });

        //var refresh = document.getElementById('refresh-clusters');
        //google.maps.event.addDomListener(refresh, 'click', refreshMap);

        //var clear = document.getElementById('clear-clusters');
        //google.maps.event.addDomListener(clear, 'click', clearClusters);

        refreshMap();
      }

      function clearClusters(e) {
        e.preventDefault();
        e.stopPropagation();
        markerClusterer.clearMarkers();
      }

      google.maps.event.addDomListener(window, 'load', initialize);
    </script>
	</body>
</html>
