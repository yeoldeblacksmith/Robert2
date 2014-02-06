<!--#include file="../classes/includelist.asp"-->
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title><%=SiteInfo.Name %> Online Waiver</title>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
        <meta name="viewport" content="width=device-width, maximum-scale=1.0, minimum-scale=1.0" /> 

        <link type="text/css" rel="stylesheet" href="content/css/waiver.css" />
        <script type="text/javascript" src="content/js/jquery-1.7.2.min.js"></script>
    </head>
    <body>
        <center>
            <br/><br/><br/>
            <table width="80%" bgcolor="ffffff" border="3">
                <tr>
                    <td>
                        <center>
                            <br/><br/><br/><br/>
                            <h1 style="text-align: center; font-size:48px">Problems Encountered</h1>
                            <h1 style="text-align: center">Your waiver was not saved.<br><br>
          
                            Please try again.</h1>
                            <br><br><br><br><br><br>
                            <%
                                If MY_SITE_GUID = "{11A9AFC7-2C12-4B4B-A8C9-375A866DB944}" Then
                                    response.write request.QueryString("src") 
                                End If
                                %>
                        </center>
                    </td>
                </tr>
            </table>
        </center>
    </body>
</html>
