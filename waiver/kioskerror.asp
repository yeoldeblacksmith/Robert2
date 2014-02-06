<!--#include file="../classes/includelist.asp"-->
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title><%=SiteInfo.Name %> Online Waiver</title>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
        <meta name="viewport" content="width=device-width, maximum-scale=1.0, minimum-scale=1.0" /> 

        <link type="text/css" rel="stylesheet" href="content/css/waiver.css" />
        <script type="text/javascript" src="content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript">
            function sendMail(waiverId) {
                $.ajax({
                    url: 'ajax/waiver.asp?act=50&id=' + waiverId,
                    error: function () { alert("problems encountered"); },
                    success: function () { alert("Waiver sent successfully."); }
                });
            }
        </script>
        
        <script type="text/JavaScript">
            setTimeout("location.href = 'kiosk.asp';", 7500);
        </script>  
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
                        </center>
                    </td>
                </tr>
            </table>
        </center>
    </body>
</html>
