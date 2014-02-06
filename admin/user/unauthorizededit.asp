<!--#include file="../../classes/IncludeList.asp" -->
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Unauthorized Profile Edit</title>

        <link type="text/css" rel="Stylesheet" href="../../content/css/eventregistration.css" />
        <link type="text/css" rel="stylesheet" href="../../content/css/admin.css" />
        <link type="text/css" rel="stylesheet" href="../../content/css/navmenu.css" />
    </head>

    <body>
<%
    dim navmenu
    set navmenu = new NavigationMenu
    navmenu.HeadingTitle = "Unauthorized Profile Edit"

    navmenu.WriteNavigationSection NAVIGATION_NAME_USERS
%>
        <p></p>

        <table class="adminTable" style="margin: 0 auto; width: 550px; height: auto;">
            <tr>
                <td class="adminCell" style=" padding: 10px">
                    <h1 class="center">Invalid Profile</h1>
                    <p>You are not authorized to edit the selected profile</p>
                </td>
            </tr>
        </table>
    </body>
</html>
