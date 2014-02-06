<!DOCTYPE html>
<!--#include file="../../classes/includelist.asp"-->
<%
    ' authenticate and authorize using pl/securityhelper
    AuthorizeForAdminOnly
%>

<html>
    <head>
        <title>Event Types</title>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">

        <link type="text/css" rel="Stylesheet" href="../../content/css/eventregistration.css" />
        <link type="text/css" rel="stylesheet" href="../../content/css/admin.css" />
        <link type="text/css" rel="stylesheet" href="../../content/css/navmenu.css" />

        <script type="text/javascript" src="../../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript">
            $(document).ready(function () {
                $.ajax({
                    url: '../../ajax/eventtype.asp?act=0',
                    success: function (data) { $("span[name=types]").html(data); },
                    error: function () { alert("problems encountered"); }
                });
            });

            function deleteType(id) {
                $.ajax({
                    url: '../../ajax/eventtype.asp?act=2&id=' + id,
                    success: function (data) { $("span[name=types]").html(data); },
                    error: function () { alert("problems encountered"); }
                });
            }

            function saveType() {
                $.ajax({
                    url: '../../ajax/eventtype.asp?act=1&ds=' + $("input[name=Description]").val(),
                    success: function (data) { $("span[name=types]").html(data); },
                    error: function () { alert("problems encountered"); }
                });
            }
        </script>
    </head>

    <body>
<%
    dim navmenu
    set navmenu = new NavigationMenu
    navmenu.HeadingTitle = "Event Type Administration"

    navmenu.WriteNavigationSection NAVIGATION_NAME_SETTINGS
%>

        <table class="adminTable" style="margin: 0 auto; width: 550px; height: auto;">
            <tr>
                <td class="adminCell" style=" padding: 10px">
                    <span name="types"></span>
                </td>
            </tr>
        </table>
    </body>
</html>

<% 

%>