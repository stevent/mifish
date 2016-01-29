<%
OPTION EXPLICIT

%>
<!--#include virtual="/app/core/core.includes.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
  <head>
    <base href="<%= APPLICATION("SiteURL") %>/<%= APPLICATION("AdminRoot") %>/" />
    <meta http-equiv="Content-type" content="text/html; charset=utf-8" />
    <title>Administration Area</title>
    <!--#include virtual="admin/public/includes/_inc-scripts_styles.asp"-->
  </head>

  <body id="p-action" class="action cms_error">
    <div id="wrapper">
      <div class="container">
        <div id="header">
          <!--#include virtual="admin/public/includes/_inc-banner.asp"-->
          <!--#include virtual="admin/public/includes/_inc-navigation.asp"-->
        </div><!--endDiv header-->

        <div id="content" class="clearfix">
          <div id="primary" class="text">
            <div class="p-head-container">
              <h1>Administration Login</h1>
            </div>

            <div class="p-text-container">
              <div class="text">
                <h2>Please use the form below to login</h2>

                <form id="edit_record" class="baseform validate_form" method="post" action="/<%= APPLICATION("AdminRoot") %>/">
                  <fieldset>
                    <legend class="hidden">Log In</legend>
                    <h2 class="title padding">Log In</h2>
                    <ul class="margin">
                      <li>
                        <label for="username">Username</label>
                        <input type="text" id="username" name="LOGIN_Username" value="" />
                      </li>
                      <li>
                        <label for="password">Password</label>
                        <input type="text" id="password" name="LOGIN_Password" value="" />
                      </li>
                      <li class="button btns_cont">
                        <label for="FormAction_Submit" class="hidden">Log In</label>
                        <button type="submit" value="Log In" name="FormAction" id="FormAction_Submit" class="btn"> Log In </button>
                      </li>
                    </ul>
                  </fieldset>
                </form>
              </div>
            </div>
          </div><!--endDiv primary-->
        </div><!--endDiv content-->

        <!--#include virtual="admin/public/includes/_inc-footer.asp"-->
      </div><!--endDiv container-->
    </div><!--endDiv wrapper-->
  </body>
</html>