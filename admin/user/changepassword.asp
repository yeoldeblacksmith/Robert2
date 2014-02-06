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
        <title>User Profile : Change Password</title>

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

                
                if ($("#Password").val().length == 0) {
                    valid = false;
                    $("#valPasswordReq").css("display", "block");
                } else if ($("#Password").val().length < 4) {
                    valid = false;
                    $("#valPasswordReq").css("display", "none");
                    $("#valPasswordMinimum").css("display", "block");
                } else {
                    $("#valPasswordReq").css("display", "none");
                    $("#valPasswordMinimum").css("display", "none");

                    if ($("#Password").val() != $("#ConfirmPassword").val()) {
                        valid = false;
                        $("#valConfirmMatch").css("display", "block");
                    } else {
                        $("#valConfirmMatch").css("display", "none");
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
                    <h1 class="center">Change Password</h1>
                    <form method="post" action="save.asp" autocomplete="off">
                        <input type="Hidden" Name="ChangePassword" id="ChangePassword" value="true" />
                        <input type="hidden" name="NoPass" value="false" />
                        <input type="hidden" name="UserName" id="UserName" value="<%= myUser.UserName %>" />

                        <table class="tableDefault marginCenter">
                            <tr>
                                <td>Password:</td>
                                <td>
                                    <input type="password" name="Password" id="Password" />
                                    <span id="valPasswordReq" class="red" style="display: none">Required</span>
                                    <span id="valPasswordMinimum" class="red" style="display: none">Password must be at least 4 characters</span>
                                </td>
                            </tr>
                            <tr>
                                <td>Confirm:</td>
                                <td>
                                    <input type="password" name="ConfirmPassword" id="ConfirmPassword" />
                                    <span id="valConfirmMatch" class="red" style="display: none">Passwords do not match</span>
                                </td>
                            </tr>
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