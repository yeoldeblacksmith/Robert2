<!DOCTYPE html>
<!--#include file="../../../classes/IncludeList.asp" -->
<%
    ' authenticate and authorize using pl/securityhelper
    AuthorizeForSiteUsers
%>
<%
    dim field
    set field = new WaiverCustomField

    if request.ServerVariables("REQUEST_METHOD") = "GET" then
        if len(Request.QueryString("id")) > 0 then
            field.WaiverCustomFieldId = request.QueryString("id")
            field.Load request.QueryString("id")
            
            if field.SiteGuid <> MY_SITE_GUID then response.Redirect("../../unauthorized.asp")
        end if
    elseif Request.ServerVariables("REQUEST_METHOD") = "POST" then
        field.WaiverCustomFieldId = request.Form("txtCustomFieldId")
        
        field.Delete

        response.Redirect "default.asp"
    end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Delete Custom Field</title>
        <link type="text/css" rel="Stylesheet" href="../../../content/css/eventregistration.css" />
        <link type="text/css" rel="Stylesheet" href="../../../content/css/admin.css" />
        <link type="text/css" rel="Stylesheet" href="../../../content/css/navmenu.css" />

        <script type="text/javascript" src="../../../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript">
        </script>
    </head>

    <body>
<%
    dim navmenu
    set navmenu = new NavigationMenu
    navmenu.HeadingTitle = "Delete Custom Field"

    navmenu.WriteNavigationSection NAVIGATION_NAME_SETTINGS 
%>
        <form method="post">
            <table class="adminTable" style="margin: 0 auto; width: 550px; height: auto;">
                <tr>
                    <td class="adminCell" style=" padding: 10px">
                        <input type="hidden" name="txtCustomFieldId" value="<%= field.WaiverCustomFieldId %>" />
                        
                        <h2 class="center">Confirm Delete</h2>

                        <p class="center">Do you really want to delete this field?</p>
                        <table class="tableDefault marginCenter">
                            <tr>
                                <td>Name:</td>
                                <td><%= field.Name %></td>
                            </tr>
                            <tr>
                                <td colspan="2" class="center">
                                    <button type="submit">Yes</button>
                                    &nbsp;
                                    <button type="button" onclick="javascript:window.location.href='default.asp';">No</button>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </form>
    </body>
</html>
