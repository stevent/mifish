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
' File:				    member-login.asp
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
<!DOCTYPE HTML>
<html lang="en">
  <head>
		<base href="<%= APPLICATION("SiteURL") %>" />
    <title>Login | Members | MiFish Online</title>

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
					<h1>Login</h1>
				</header>
				<div class="box">
					<form id="login" action="<%= APPLICATION("SiteURL") %>/member-home.asp" method="post">
						<div class="row uniform 50%">
							<div class="12u">
								<input type="text" name="username" id="login_username" placeholder="Username" value="" />
							</div>
						</div>
						<div class="row uniform 50%">
							<div class="12u">
								<input type="password" name="password" id="login_password" placeholder="Password" value="" />
							</div>
						</div>
						<div class="row uniform">
							<div class="12u">
								<ul class="actions align-center">
									<li><input type="submit" value="Login" /></li>
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
