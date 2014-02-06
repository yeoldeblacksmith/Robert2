<!--#include file="../classes/includelist.asp"-->
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Vantora Online Waiver</title>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
        <meta name="viewport" content="width=device-width, maximum-scale=1.0, minimum-scale=1.0" /> 

        <link rel="stylesheet" href="../content/css/jquery.signaturepad.css" />

        <!--[if lt IE 9]><script src="../content/js/flashcanvas.js"></script><![endif]-->
        <script src="../content/js/jquery-1.7.2.min.js"></script>
        <script src="../content/js/jquery.signaturepad.min.js"></script>
        <script src="../content/js/jquery.signaturepad.compression.js"></script>
        <script src="../content/js/json2.min.js"></script>
    </head>

    <body>
        <h1 style="text-align: center">Email Sent</h1>
        <p>Thank you for taking the time to sign your insurance waiver online.</p>
        <p>Next Steps:</p>
        <ul>
            <li><a href="display.asp?id=<%= Request.QueryString(QUERYSTRING_VAR_WAIVERID) %>">View printable version of this waiver</a></li>
            <li><a href="sendemail.asp?id=<%= Request.QueryString(QUERYSTRING_VAR_WAIVERID) %>">Email me a link to this waiver</a></li>
            <li><a href="default.asp?id=<%= Request.QueryString(QUERYSTRING_VAR_WAIVERID) %>">Fill out another waiver</a></li>
            <li><a href="<%= SiteInfo.HOmeUrl %>">Go home</a></li>
        </ul>
    </body>
</html>
