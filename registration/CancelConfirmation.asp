<!DOCTYPE html>
<!--#include file="../classes/IncludeList.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Registration Cancelled</title>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
        <meta name="viewport" content="width=device-width, maximum-scale=1.0, minimum-scale=1.0" /> 

        <link type="text/css" rel="stylesheet" href="../content/css/jquery.mobile-1.1.0.css" />
        <link type="text/css" rel="Stylesheet" href="../content/css/eventregistration.css" />
        <link type="text/css" rel="Stylesheet" href="../content/css/waiver.css" />
    </head>

    <body>
        <div class="waiverSection blockCenter" style="width: 550px">
            <img src="<%= LOGO_URL %>" style="display: block; margin: auto" />
            <h1 style="text-align:center">Registration Cancelled</h1>

            <p style="text-align:center">
                You have successfully canceled your registration.
                <br /><br />
                <a href="default.asp">Register for another event</a>
                <br /><br />
                <a href="<%= SiteInfo.HomeUrl %>">Return home</a>
            </p>
        </div>
    </body>
</html>
