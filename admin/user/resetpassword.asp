<!DOCTYPE HTML>
<%
    dim userHash
    userHash = request.QueryString("u")

    if len(userHash) > 0 then
        userHash = request.QueryString("u")
    else
        SendToLogin()
    end if
%>
<!--#include file="../../classes/IncludeList.asp" -->
<%
    dim attempCount, myUser

    set myUser = new User
    myUser.PasswordResetHash = userHash
    myUser.LoadByUserHash()

    if len(myUser.Password) <= 0 then
        SendToLogin()
    end if

    if len(request.Form("atmp")) > 0 then attemptCount = cint(request.Form("atmp")) else attemptCount = 0

    if len(request.Cookies("role")) > 0 then
        RedirectUser
    elseif Request.ServerVariables("REQUEST_METHOD") = "POST" then
        myUser.Password = Request.Form("Password")
        myUser.Save
        
        Response.Cookies("site") = myUser.SiteGuid
        Response.Cookies("role") = myUser.Role.RoleId
        Response.Cookies("user") = myUser.UserName
        RedirectUser
    end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>Admin Login</title>
    
    <script type="text/javascript" src="../../content/js/jquery-1.7.2.min.js"></script>
    <script type="text/javascript">
        $(document).ready(function () {
            $("[name=Password]").focus();
        });

        $("#cancel").click(function() {
                window.location='../signout.asp';
        });

        function validateForm() {
            var valid = true;

            
            if ($("#Password").val().length == 0) {
                valid = false;
                $("#valPasswordReq").css("display", "inline");
            } else if ($("#Password").val().length < 4) {
                valid = false;
                $("#valPasswordReq").css("display", "none");
                $("#valPasswordMinimum").css("display", "inline");
            } else {
                $("#valPasswordReq").css("display", "none");
                $("#valPasswordMinimum").css("display", "none");

                if ($("#Password").val() != $("#ConfirmPassword").val()) {
                    valid = false;
                    $("#valConfirmMatch").css("display", "inline");
                } else {
                    $("#valConfirmMatch").css("display", "none");
                }
            }
            return valid;
        }
            
    </script>
</head>

<body>
    <br><br><center><table width="500" border="3" cellpadding="10" bgcolor="#eeeeee"><tr><td>
        <center><img src="<%= LOGO_URL %>"><br><h3>Reset your Password</h3><p>

        <form action='<%Response.Write "resetpassword.asp?"&Request.QueryString %>' method="post">
            <input type="hidden" name="PasswordReset" value="true" />
            <input type="hidden" name="User" value='<%= myUser.UserName %>' />
            <table width="100%">
                <tr>
                    <td width="30%" align="right">
                        Password:
                    </td>
                    <td width="50%" align="left">
                        <input type="password" id="Password" name="Password" value='' size="20"/>
                    </td>
                </tr>
                <tr>
                    <td>
                        &nbsp;
                    </td>
                    <td>
                        <span id="valPasswordReq" class="red" style="display: none; color: red;">Required</span>
                        <span id="valPasswordMinimum" class="red" style="display: none; color: red;">Password must be at least 4 characters</span>
                    </td>
                </tr>
                <tr>
                    <td width="30%" align="right">
                        Confirm:
                    </td>
                    <td align="left">
                        <input type="password" id="ConfirmPassword" name="ConfirmPassword" value='' size="20"/>
                    </td>
                </tr>
                <tr>
                    <td>
                        &nbsp;
                    </td>
                    <td>
                        <span id="valConfirmMatch" class="red" style="display: none; color: red;">Passwords do not match</span>
                    </td>
                </tr>
                <tr>
                    <td>
                        &nbsp;
                    </td>
                    <td>
                        <input type="submit" value="Change" onclick="return validateForm()"/>
                        <input id="cancel" type="button" value="Cancel"/>
                    </td>
                </tr>
            </table>
        </form>
        </p></center>
    </td></tr></table></center>
</body>
</html>
<%
    sub RedirectUser()
        'Check where the users are coming from within the application.
        If (Request.QueryString("from")<>"") then
	        Response.Redirect Request.QueryString("from")
        else
	    'If the first page that the user accessed is the Logon page,
            'direct them to the default page.
            if cbool(settings(SETTING_MODULE_REGISTRATION)) then
                Response.Redirect "../events/default.asp"
            else
                Response.Redirect "../waiver/default.asp"
            end if
        End if    
    end sub

    sub SendToLogin()
        Response.Redirect "../default.asp"
    end sub
%>