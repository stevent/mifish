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
' File:				    member-waypoints-import.asp
' Version:			  1.0
' Author:			    Steven Taddei
' Modified By:		Steven Taddei
' Date Created:	  18/01/2016
' Last Modified:	18/01/2016
'
' Purpose:
'	select what sounder you want to import from
'***************************************
%>
<!--#include virtual="/admin/core/core.includes.asp"-->
<%
CALL Member.run("Login",NULL)
CALL Member.run("SecurePage",NULL)
DIM oSounders    : SET oSounders  = c_Sounder.run("FindByMember",Member.FieldValue("ID"))
DIM oSounder
%>
<!DOCTYPE HTML>
<html lang="en">
  <head>
		<base href="<%= APPLICATION("SiteURL") %>" />
    <title>Import Waypoints | Members | MiFish Online</title>

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
					<h1>Import Waypoints</h1>
				</header>
				<div class="box">
          <form id="waypoints" action="<%= APPLICATION("SiteURL") %>/member-waypoints-action.asp" method="post" enctype="multipart/form-data">
            <div class="hidden_fields">
              <input type="text" name="Action" id="waypoints_action" value="import" />
            </div>
						<div class="row uniform 50%">
              <!-- Type, Title, Longitude, Latitude, Notes -->
							<div class="12u">
                <label>What Import</label>
                <select name="SounderID" id="waypoints_sounderid">
                  <option value="">Select Type</option>
                  <option value="DDMMSS">CSV DDMMSS</option>
                  <option value="DDMM.MMM">CSV DDMM.MMM</option>
<%
FOR EACH oSounder IN oSounders.ITEMS
%>
                  <option value="<%= oSounder.FieldValue("ID") %>"><%= oSounder.FieldValue("Name") %></option>
<%
NEXT
%>
                </select>
							</div>
						</div>
            <div class="row uniform 50%">
							<div class="12u">
                <label>File</label>
                <input type="file" name="ImportFile" id="waypoints_importfile" />
							</div>
						</div>
						<div class="row uniform">
							<div class="12u">
								<ul class="actions align-center">
									<li><input type="submit" value="Import Waypoints" class="disable-on-click" /></li>
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
