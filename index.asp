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
' File:				    index.asp
' Version:			  1.0
' Author:			    Steven Taddei
' Modified By:		Steven Taddei
' Date Created:	  12/01/2016
' Last Modified:	12/01/2016
'
' Purpose:
'	Shows page text for the home page as well
' as other dynamic and static homepage
' content
'***************************************
%>
<!--#include virtual="/admin/core/core.includes.asp"-->
<%
DIM oPage : SET oPage = c_Page.run("FindByID",1)
%>
<!DOCTYPE HTML>
<html lang="en">
  <head>
		<base href="<%= APPLICATION("SiteURL") %>" />
    <title><%= oPage.FieldValue("PageTitle") %></title>

		<!--#include virtual="/assets/views/scripts-styles.asp"-->
	</head>
	<body class="landing">
		<div id="page-wrapper">
			<!-- Header -->
			<header id="header" class="alt">
				<!--#include virtual="/assets/views/header.asp"-->
			</header>

			<!-- Banner -->
			<section id="banner">
				<div class="banner_title">MiFish Online</div>
				<p>Your online fishing resource</p>
				<ul class="actions">
<%
'have we got someone logged in'
IF ( Member.run("isLoggedIn",FALSE) ) THEN
%>
          <li><a href="member-home.asp" class="button special">Your Home</a></li>
          <li><a href="member-waypoints.asp" class="button">Your Waypoints</a></li>
<%
ELSE
%>
          <!-- <li><a href="member-signup.asp" class="button special">Sign Up</a></li> -->
          <li><a href="member-login.asp" class="button">Login</a></li>
<%
END IF
%>

				</ul>
			</section>

			<!-- Main -->
			<section id="main" class="container">
				<section class="box special">
					<header class="major">
						<h1><%= oPage.run("getHeading",FALSE) %></h1>
						<p>Blandit varius ut praesent nascetur eu penatibus nisi risus faucibus nunc ornare<br />
						adipiscing nunc adipiscing. Condimentum turpis massa.</p>
					</header>
					<!-- <span class="image featured"><img src="assets/images/pic01.jpg" alt="" /></span> -->
				</section>
<!--
				<section class="box special features">
					<div class="features-row">
						<section>
							<span class="icon major fa-bolt accent2"></span>
							<h3>Magna etiam</h3>
							<p>Integer volutpat ante et accumsan commophasellus sed aliquam feugiat lorem aliquet ut enim rutrum phasellus iaculis accumsan dolore magna aliquam veroeros.</p>
						</section>
						<section>
							<span class="icon major fa-area-chart accent3"></span>
							<h3>Ipsum dolor</h3>
							<p>Integer volutpat ante et accumsan commophasellus sed aliquam feugiat lorem aliquet ut enim rutrum phasellus iaculis accumsan dolore magna aliquam veroeros.</p>
						</section>
					</div>
					<div class="features-row">
						<section>
							<span class="icon major fa-cloud accent4"></span>
							<h3>Sed feugiat</h3>
							<p>Integer volutpat ante et accumsan commophasellus sed aliquam feugiat lorem aliquet ut enim rutrum phasellus iaculis accumsan dolore magna aliquam veroeros.</p>
						</section>
						<section>
							<span class="icon major fa-lock accent5"></span>
							<h3>Enim phasellus</h3>
							<p>Integer volutpat ante et accumsan commophasellus sed aliquam feugiat lorem aliquet ut enim rutrum phasellus iaculis accumsan dolore magna aliquam veroeros.</p>
						</section>
					</div>
				</section>

				<div class="row">
					<div class="6u 12u(narrower)">

						<section class="box special">
							<span class="image featured"><img src="assets/images/pic02.jpg" alt="" /></span>
							<h3>Sed lorem adipiscing</h3>
							<p>Integer volutpat ante et accumsan commophasellus sed aliquam feugiat lorem aliquet ut enim rutrum phasellus iaculis accumsan dolore magna aliquam veroeros.</p>
							<ul class="actions">
								<li><a href="#" class="button alt">Learn More</a></li>
							</ul>
						</section>

					</div>
					<div class="6u 12u(narrower)">

						<section class="box special">
							<span class="image featured"><img src="assets/images/pic03.jpg" alt="" /></span>
							<h3>Accumsan integer</h3>
							<p>Integer volutpat ante et accumsan commophasellus sed aliquam feugiat lorem aliquet ut enim rutrum phasellus iaculis accumsan dolore magna aliquam veroeros.</p>
							<ul class="actions">
								<li><a href="#" class="button alt">Learn More</a></li>
							</ul>
						</section>
					</div>
				</div>
			</section>
-->
			<!-- Footer -->
			<footer id="footer">
				<!--#include virtual="/assets/views/footer.asp"-->
			</footer>
		</div>

		<!--#include virtual="/assets/views/footer-scripts.asp"-->
	</body>
</html>
