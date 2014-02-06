<!DOCTYPE html>
<!--#include file="../../classes/includelist.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title></title>
        <script type="text/javascript" src="../../content/js/jquery-1.7.2.min.js"></script>

        <script type="text/javascript">
            $(document).ready(function () {
                $.ajax({
                    url: '../../ajax/waiver.asp?act=64&ev=<%= Request.QueryString(QUERYSTRING_VAR_EVENTID) %>&dt=<%= Request.Querystring(QUERYSTRING_VAR_SELECTEDDATE)%>&pt=<%= Request.Querystring(QUERYSTRING_VAR_PLAYTIME)%>',
                    error: function () { alert("problems encountered"); },
                    success: function (data) {
                        $("#content").html(data);
                    }
                });
            });
        </script>
    </head>

    <body>
        <div id="content" />
    </body>
</html>
