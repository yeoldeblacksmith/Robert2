<!DOCTYPE HTML>
<!--#include file="../classes/IncludeList.asp" -->
<%

    if len(request.Cookies("role")) > 0 then
            RedirectUser
    end if

    dim attempCount

    if len(request.Form("atmp")) > 0 then attemptCount = cint(request.Form("atmp")) else attemptCount = 0
    if len(request.Form("passwordatmp")) > 0 then passwordAttemptCount = cint(Request.Form("passwordatmp")) else passwordAttemptCount = 0

    if Login = False and PasswordExpired = False then Login = True

    if Request.ServerVariables("REQUEST_METHOD") = "POST" then
        dim myUser
        set myUser = new User
        myUser.UserName = Request.Form("User")
        myUser.Load

        attemptCount = attemptCount + 1

        if Request.Form("Login") = "true" and myUser.AuthenticateUser(Request.Form("User"), Request.Form("password")) then
            Response.Cookies("site") = myUser.SiteGuid
            Response.Cookies("role") = myUser.Role.RoleId
            Response.Cookies("user") = myUser.UserName
            RedirectUser
        end if

    end if

%>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Admin Login</title>
    
        <script type="text/javascript" src="../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript">
            $(document).ready(function () {
                $("[name=User]").focus();
            });
        </script>
    </head>

    <body>
        <br/><br/>
        <center>
            <table width="500" border="3" cellpadding="10" bgcolor="#eeeeee">
                <tr>
                    <td>
                        <center>
                            <img src="<%= LOGO_URL %>" />
                            <h3>Logon to Reservation Admin Screen</h3>

                        <% if attemptCount > 1 then %>
                            <h3 style="color: red">Authorization Failed.</h3>
                            <p style="color: red">Please try again.</p>
                        <% end if %>

                            <form action='<%Response.Write "default.asp?"&Request.QueryString%>' method="post">
                                <input type="hidden" name="Login" value="true" />
                                <input type="hidden" name="atmp" value="<%= attemptCount %>" />
                                Username: 
                                <input type="text" name="User" value='' size="20"/><br><br>
                                Password: 
                                <input type="password" name="password" value='' size="20"/><br><br>
                                <a href="forgotpassword.asp">Forgot your password?</a><br/><br/>
                                <input type="submit" value="Logon"/>

                            </form>
                        </center>
                    </td>
                </tr>
            </table>
        </center>
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
                Response.Redirect "events/default.asp"
            else
                Response.Redirect "waiver/default.asp"
            end if
        End if    
    end sub    
%>