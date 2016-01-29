        <div class="logo_text"><a href="./">MiFish Online</a></div>
        <nav id="nav">
          <ul>
            <li><a href="./">Home</a></li>
<%
'have we got someone logged in'
IF ( Member.run("isLoggedIn",FALSE) ) THEN
%>
            <li>
              <a href="member-home.asp" class="icon fa-angle-down">My Home</a>
              <ul>
                <li>
                  <a href="member-waypoints.asp">Waypoints</a>
                  <ul>
                    <li><a href="member-waypoints.asp">View Waypoints</a></li>
                    <li><a href="member-waypoints-form.asp">Add Waypoint</a></li>
                    <li><a href="member-waypoints-import.asp">Import Waypoints</a></li>
                    <li><a href="member-waypoints-export.asp">Export Waypoints</a></li>
                  </ul>
                </li>
                <li><a href="#">My Catches</a></li>
              </ul>
            </li>
            <li><a href="member-logout.asp" class="button">Log Out</a></li>
<%
ELSE
%>
            <li><a href="member-login.asp" class="button">Login</a></li>
<%
END IF
%>

          </ul>
        </nav>
