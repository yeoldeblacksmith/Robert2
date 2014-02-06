<!DOCTYPE html>
<!--#include file="../classes/IncludeList.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Unauthorized Access</title>
        <link type="text/css" rel="stylesheet" href="../content/css/admin.css" />
        <link type="text/css" rel="stylesheet" href="../content/css/navmenu.css" />
    </head>

    <body>
<%
    dim navmenu
    set navmenu = new NavigationMenu
    navmenu.HeadingTitle = "Unauthorized Access"

    navmenu.WriteNavigationSection "" 
%>
        <p></p>

        <table class="adminTable" style="margin: 0 auto; width: 550px; height: auto;">
            <tr>
                <td class="adminCell" style=" padding: 10px">
                    <h1 class="center">Unauthorized Access</h1>
                    <p class="center">You do not have access to view the requested page.</p>
                </td>
            </tr>
        </table>
    </body>
</html>
