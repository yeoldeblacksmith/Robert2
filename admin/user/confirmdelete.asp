<!DOCTYPE html>
<!--#include file="../../classes/IncludeList.asp" -->
<%
    ' authenticate and authorize using pl/securityhelper
    AuthorizeForAdminOnly

    dim myUser
    set myUser = new User

    if len(request.QueryString(QUERYSTRING_VAR_USERNAME)) > 0 then
        myUser.UserName = Request.QueryString(QUERYSTRING_VAR_USERNAME)
        myUser.Load

        myUser.Role.Load

        dim isGlobalAdmin
        isGlobalAdmin = cbool(Request.Cookies(COOKIE_ROLEID) = AUTHORIZED_ROLE_GLOBALADMIN)

        if isGlobalAdmin then
        else
            ' do not let non-global admins delete global admins
            if cstr(myUser.Role.RoleId) = AUTHORIZED_ROLE_GLOBALADMIN then 
                response.Redirect("unauthorizededit.asp")
            ' do not let users delete users from outside of their site
            elseif myUser.SiteGuid <> MY_SITE_GUID then
                response.Redirect("unauthorizededit.asp")
            end if    
        end if
        
    end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Confirm Delete</title>
        <link type="text/css" rel="Stylesheet" href="../../content/css/eventregistration.css" />
        <link type="text/css" rel="stylesheet" href="../../content/css/admin.css" />
        <link type="text/css" rel="stylesheet" href="../../content/css/navmenu.css" />
        <script type="text/javascript">
            function redirect() {
                window.location.href = "default.asp";
                return false;
            }
        </script>
    </head>

    <body>
<%
    dim navmenu
    set navmenu = new NavigationMenu
    navmenu.HeadingTitle = "Delete User Profile"

    navmenu.WriteNavigationSection NAVIGATION_NAME_USERS
%>
        <p></p>

        <table class="adminTable" style="margin: 0 auto; width: 550px; height: auto;">
            <tr>
                <td class="adminCell" style=" padding: 10px">
                    <h1 class="center">Confirm Delete</h1>
                    <p class="center">Are you sure you want to delete this user?</p>

                    <form action="delete.asp" method="post">
                        <input type="hidden" name="UserName" id="UserName" value="<%= myUser.UserName %>" />
                        <table class="tableDefault marginCenter">
                            <tr>
                                <td>User Name:</td>
                                <td><%= myUser.UserName %></td>
                            </tr>
                            <tr>
                                <td>Email Address:</td>
                                <td><%= myUser.EmailAddress %></td>
                            </tr>
                            <tr>
                                <td>Role:</td>
                                <td><%= myUser.Role.Description %></td>
                            </tr>
                            <tr>
                                <td>Enabled:</td>
                                <td>
                                    <input type="checkbox" name="Enabled" id="Enabled" checked="<%= iif(myUser.Enabled, "checked", "") %>" disabled="disabled" />
                                </td>
                            </tr>
                            <tr>
                                <td>Password Expired:</td>
                                <td>
                                    <input type="checkbox" name="PasswordExpired" id="PasswordExpired" checked="<%= iif(myUser.PasswordExpired, "checked", "") %>" disabled="disabled" />
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2" class="center">
                                    <button type="submit">Delete</button>&nbsp;
                                    <button onclick="return redirect()">Cancel</button>
                                </td>
                            </tr>
                        </table>
                    </form>

                </td>
            </tr>
        </table>
    </body>
</html>
