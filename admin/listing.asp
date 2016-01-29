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

  <body id="p-listing" class="listing <%= LCASE(g_CurrentController) %>">
    <div id="wrapper">
      <div class="container">
        <div id="header">
          <!--#include virtual="admin/public/includes/_inc-banner.asp"-->
          <!--#include virtual="admin/public/includes/_inc-navigation.asp"-->
        </div><!--endDiv header-->

        <div id="content" class="clearfix">
          <div id="primary" class="text">
            <div class="p-head-container">
              <h1>Listing Title</h1>
            </div>

            <div class="p-text-container">
              <div class="text">

                <table id="section-list" class="listing-table">
                  <thead>
                    <tr>
                      <th>list name</th>
                      <th>list name</th>
                      <th>list name</th>
                      <th>actions</th>
                    </tr>
                  </thead>

                  <tbody>
                    <tr class="hover">
                      <td>list info</td>
                      <td>list info</td>
                      <td>list info</td>
                      <td>
                        <a href="#">edit</a> <a href="#">delete</a>
                      </td>
                    </tr>
                    <tr class="hover">
                      <td>list info</td>
                      <td>list info</td>
                      <td>list info</td>
                      <td>
                        <a href="#">edit</a> <a href="#">delete</a>
                      </td>
                    </tr>
                    <tr class="hover">
                      <td>list info</td>
                      <td>list info</td>
                      <td>list info</td>
                      <td>
                        <a href="#">edit</a> <a href="#">delete</a>
                      </td>
                    </tr>
                    <tr class="hover">
                      <td>list info</td>
                      <td>list info</td>
                      <td>list info</td>
                      <td>
                        <a href="#">edit</a> <a href="#">delete</a>
                      </td>
                    </tr>
                  </tbody>
                </table>

                <div class="btns_cont clearfix">
                  <div class="new_link">
                    <a href="#" class="btn" title="new record">new</a>
                  </div>
                </div>
              </div>
            </div>
          </div><!--endDiv primary-->
        </div><!--endDiv content-->

        <!--#include virtual="admin/public/includes/_inc-footer.asp"-->
      </div><!--endDiv container-->
    </div><!--endDiv wrapper-->
  </body>
</html>