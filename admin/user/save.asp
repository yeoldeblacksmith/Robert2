<!--#include file="../../classes/IncludeList.asp" -->

<%
    dim myUser
    set myUser = new User


    myUser.UserName = request.Form("UserName")    

    if len(request.Form("ChangePassword")) = 0 then
        myUser.EmailAddress = request.Form("EmailAddress")
        myUser.Role.RoleId = request.Form("RoleId")
        myUser.Enabled = iif(request.Form("Enabled") = "on", true, false)
        myUser.PasswordExpired = iif(request.Form("PasswordExpired") = "on", true, false)
        myUser.Save
    end if

    if (request.form("NoPass") = "false") then
        myUser.Password = request.Form("Password")
        myUser.ChangePassword
    end if

    'myUser.PasswordExpired = false

    response.Redirect("default.asp")
%>