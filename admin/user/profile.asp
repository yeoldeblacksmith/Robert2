<!DOCTYPE html>
<!--#include file="../../classes/IncludeList.asp" -->
<%
    ' authenticate and authorize using pl/securityhelper
    AuthorizeForAdminOnly

    dim myUser, editMode, globalAdmin
    set myUser = new User
    editMode = false
    globalAdmin = cbool(request.Cookies("role") = AUTHORIZED_ROLE_GLOBALADMIN)

    if len(request.QueryString(QUERYSTRING_VAR_USERNAME)) > 0 then
        myUser.UserName = Request.QueryString(QUERYSTRING_VAR_USERNAME)
        myUser.Load

        if cstr(myUser.Role.RoleId) = AUTHORIZED_ROLE_GLOBALADMIN and globalAdmin = false then response.Redirect("unauthorizededit.asp")

        editMode = true
    end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>User Profile</title>

        <link type="text/css" rel="Stylesheet" href="../../content/css/eventregistration.css" />
        <link type="text/css" rel="stylesheet" href="../../content/css/admin.css" />
        <link type="text/css" rel="stylesheet" href="../../content/css/navmenu.css" />

        <script type="text/javascript" src="../../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript">
            function redirect() {
                window.location.href = "default.asp";
                return false;
            }

            function validateForm() {
                var valid = true;

                if ($("#UserName").val().length == 0) {
                    valid = false;
                    $("#valUserNameReq").css("display", "block");
                } else {
                    $("#valUserNameReq").css("display", "none");
                }
                
                var email = $("#EmailAddress");
                if (email.val().length === 0) {
                    valid = false;
                    $("#valEmailAddressReq").css("display", "block");
                } else {
                    $("#valEmailAddressReq").css("display", "none");
                    var emailRegExp = /^([a-zA-Z0-9_\.\-\+])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9]{2,4})+$/;
                    if (emailRegExp.test(email.val()) == false) {
                        valid = false;
                        $("#valEmailAddressFormat").css("display", "block");
                    } else {
                        $("#valEmailAddressFormat").css("display", "none");

                        if (email.val() != $("#ConfirmAddress").val()) {
                            valid = false;
                            $("#valConfirmAddressMatch").css("display", "block");
                        } else {
                            $("#valConfirmAddressMatch").css("display", "none");
                        }
                    }
                }
                return valid;
            }
        </script>
    </head>

    <body>
<%
    dim navmenu
    set navmenu = new NavigationMenu
    navmenu.HeadingTitle = "User Profile"

    navmenu.WriteNavigationSection NAVIGATION_NAME_USERS
%>
        <p></p>

        <table class="adminTable" style="margin: 0 auto; width: 550px; height: auto;">
            <tr>
                <td class="adminCell" style=" padding: 10px">
                    <h1 class="center">User Profile</h1>
                    <form method="post" action="save.asp" autocomplete="off">

                        <table class="tableDefault marginCenter">
                            <tr>
                                <td>User Name:</td>
                                <td>
                                    <input type="text" name="UserName" id="UserName" value="<%= myUser.UserName %>" <%= iif(editMode, "readonly=""readonly""", "") %> /><br />
                                    <span id="valUserNameReq" class="red" style="display: none">Required</span>
                                </td>
                            </tr>
                            <tr>
                                <td>Password:</td>
                                <td>
                                    <a href="changepassword.asp?un=<%= myUser.UserName %>">change password</a><br />
                                    <input type="hidden" name="NoPass" value="true" />
                                </td>
                            </tr>
                            <tr>
                                <td>Email Address:</td>
                                <td>
                                    <input type="text" name="EmailAddress" id="EmailAddress" value="<%= myUser.EmailAddress %>" /><br />
                                    <span id="valEmailAddressReq" class="red" style="display: none">Required</span>
                                    <span id="valEmailAddressFormat" class="red" style="display: none">Invalid email address</span>
                                </td>
                            </tr>
                            <tr>
                                <td>Confirm:</td>
                                <td>
                                    <input type="text" name="ConfirmAddress" id="ConfirmAddress" value="<%= myUser.EmailAddress %>" /><br />
                                    <span id="valConfirmAddressMatch" class="red" style="display: none">Email addresses do not match</span>
                                </td>
                            </tr>
                            <tr>
                                <td>Role:</td>
                                <td>
                                    <% BuildRoleList myUser.Role.RoleId %>
                                </td>
                            </tr>
                            <tr>
                                <td>Enabled:</td>
                                <td>
                                    <input type="checkbox" name="Enabled" id="Enabled" checked="<%= iif(myUser.Enabled, "checked", "") %>" />
                                </td>
                            </tr>
<!--                            <tr>
                                <td>Password Expired:</td>
                                <td>
                                    <input type="checkbox" name="PasswordExpired" id="PasswordExpired" checked="<%= iif(myUser.PasswordExpired, "checked", "") %>" />
                                </td>
                            </tr>-->
                            <tr>
                                <td colspan="2" class="center">
                                    <button type="submit" onclick="return validateForm()">Update</button>&nbsp;
                                    <button onclick="return redirect()">Cancel</button>
                                </td>
                            </tr>
                        </table>

                    </form>
                </td>
            </tr>
        </table>
    </body>
</html>
<% 
sub BuildRoleList(SelectedValue)
    dim roles, startIndex
    set roles = new AuthorizationRoleCollection
    roles.Load

    response.Write "<select name=""RoleId"" id=""RoleId"">" & vbCrLf

    if globalAdmin then startIndex = 0 else startIndex = 1

    for i = startIndex to roles.count - 1
        response.Write "<option value=""" & roles(i).RoleId & """ " & iif(roles(i).RoleId = SelectedValue, "selected","") & ">" & roles(i).Description & "</option>" & vbCrLf
    next

    response.Write "</select>" & vbCrLf
end sub    
%>