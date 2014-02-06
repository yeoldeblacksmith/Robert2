<!DOCTYPE html>
<!--#include file="../../classes/IncludeList.asp" -->
<%
    ' authenticate and authorize using pl/securityhelper
    AuthorizeForAdminOnly
%>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>What's New</title>
        
        <link type="text/css" rel="Stylesheet" href="../../content/css/eventregistration.css" />
        <link type="text/css" rel="stylesheet" href="../../content/css/admin.css" />
        <link type="text/css" rel="stylesheet" href="../../content/css/navmenu.css" />
    </head>

    <body>
<%
    dim navmenu
    set navmenu = new NavigationMenu
    navmenu.HeadingTitle = "What's New"

    navmenu.WriteNavigationSection ""
%>
        <table class="adminTable" style="margin: 0 auto; width: 1050px; height: auto;">
            <tr>
                <td class="adminCell" style=" padding: 10px">
                    <table class="tableDefault marginCenter" style="border-spacing: 5px">
                        <tr>
                        <td>

 <h2>Instructions</h2>
 <h3>Getting Started</h3>
 The first thing you need to do is to...
                                            
                                            
                             </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </body>
</html>
