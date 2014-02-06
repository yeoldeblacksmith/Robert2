<!DOCTYPE HTML>
<!--#include file="../classes/IncludeList.asp" -->
<%
    dim attempCount, successful
    if len(request.Form("atmp")) > 0 then attemptCount = cint(request.Form("atmp")) else attemptCount = 0
    successful = false

    if len(request.Cookies("role")) > 0 then
        RedirectUser
    elseif Request.ServerVariables("REQUEST_METHOD") = "POST" then
        dim myUser
        set myUser = new User

        myUser.LoadByEmail(Request.Form("Email"))

        if len(myUser.UserName) > 0 then
            myUser.SaveResetPasswordHash myUser.UserName, MY_SITE_GUID
            successful = true
        else
            attemptCount = attemptCount + 1
        end if
    
    end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Admin Login</title>
    
        <script type="text/javascript" src="../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript">
            $(document).ready(function () {
                $("[name=Email]").focus();
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
                            <img src="<%= LOGO_URL %>"/>
                            <h3>Reset your password</h3>

                        <% if attemptCount >= 1 And attemptCount < 3 And successful = false then %>
                            <h3 style="color: red">User not found.</h3>
                            <p style="color: red">Please try again.</p>
                        <% elseif attemptCount >= 3 And successful = false then %>
                            <h3 style="color: red">User not found.</h3>
                            <p style="color: red">Please contact support if you're having trouble remembering your username.</p>
                        <% end if %>

                        <% if successful = false then %>
                            <form action='<%Response.Write "forgotpassword.asp?"&Request.QueryString%>' method="post">
                                <input type="hidden" name="atmp" value="<%= attemptCount %>" />
                                Email: 
                                <input type="text" name="Email" value='' size="20"/><br><br>
                                <input type="submit" value="Reset Password"/>
                            </form>
                        <% else %>
                            <h3>An email with reset instructions has been sent.</h3>
                            <a href='<%= PathCombine(SiteInfo.VantoraUrl, "admin") %>'>Return to login.</a>
                        <% end if %>
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