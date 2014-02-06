<!DOCTYPE html>
<!--#include file="../../classes/pureaspupload.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title></title>

        <script type="text/javascript" src="../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/JavaScript">
            $(document).ready(function () { $('#doingupload').css("display", "none"); });

            function onSubmit() {
                $('#beforeupload').css("display", "none");
                $('#doingupload').css("display", "table-row");

                return false;
            }
        </script>
    </head>

    <body>
        <form method="post" enctype="multipart/form-data" onsubmit="return onSubmit();">
          <table style="width: 100%">
  	        <tr id="beforeupload">
  		        <td style="width: 300px">
                    <input type="file" name="uploadFile" id="uploadFile" style="width: 100%"/>
                </td>
  		        <td>
                    <button type="submit">Upload</button>
                </td>
  	        </tr>
		    <tr id="doingupload">
			    <td colspan="2" style="text-align: center">
				    Please wait for your file to be uploaded...<br />
				    <img alt="" id="waitingpic" src="../../content/images/loader.white.gif" style="text-align:center;vertical-align:middle;" />
			    </td>
		    </tr>
          </table>
        </form>
    </body>
</html>
