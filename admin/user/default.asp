<!DOCTYPE html>
<!--#include file="../../classes/IncludeList.asp" -->
<%
    ' authenticate and authorize using pl/securityhelper
    AuthorizeForAdminOnly
%>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>User Administration</title>
        
        <link type="text/css" rel="Stylesheet" href="../../content/css/eventregistration.css" />
        <link type="text/css" rel="stylesheet" href="../../content/css/admin.css" />
        <link type="text/css" rel="stylesheet" href="../../content/css/navmenu.css" />
    </head>

    <body>
<%
    dim navmenu
    set navmenu = new NavigationMenu
    navmenu.HeadingTitle = "User Administration"

    navmenu.WriteNavigationSection NAVIGATION_NAME_USERS
%>
        <p></p>

        <table class="adminTable" style="margin: 0 auto; width: 750px; height: auto;">
            <tr>
                <td class="adminCell" style=" padding: 10px">
                    <table class="tableDefault marginCenter" style="border-spacing: 5px">
                        <tr>
                            <th>User Name</th>
                            <th>Email Address</th>
                            <th>Role</th>
                            <th>Enabled</th>
                            <th colspan="2" />
                        </tr>

<% 
    dim myUsers
    set myUsers = new UserCollection
    myUsers.LoadForSite

    for i = 0 to myUsers.Count - 1
        myUsers(i).Role.Load

        ' do not show global admins to site admins
        dim viewableRecord : viewableRecord = false
        if Request.Cookies(COOKIE_ROLEID) = AUTHORIZED_ROLE_GLOBALADMIN then
            viewableRecord = true
        else
            if myUsers(i).Role.RoleId <> cint(AUTHORIZED_ROLE_GLOBALADMIN) then
                viewableRecord = true
            end if
        end if

        if viewableRecord then
            response.Write "<tr>" & vbCrLf
        
            response.Write "<td>" & myUsers(i).UserName & "</td>" & vbCrLf
            response.Write "<td>" & myUsers(i).EmailAddress & "</td>" & vbCrLf
            response.Write "<td>" & myUsers(i).Role.Description & "</td>" & vbCrLf
            response.Write "<td class=""center""><input type=""checkbox"" disabled=""disabled"" " & iif(myUsers(i).Enabled, "checked=""checked""", "")  & "/></td>" & vbCrLf
        
            dim showLinks
            showLinks = false    

            select case CStr(myUsers(i).Role.RoleId)
                case AUTHORIZED_ROLE_GLOBALADMIN
                    if Request.Cookies("role") = AUTHORIZED_ROLE_GLOBALADMIN then showLinks = true
                case AUTHORIZED_ROLE_ADMIN, AUTHORIZED_ROLE_MANAGER
                    showLinks = true
            end select

            if showLinks then
                response.Write "<td><a href=""profile.asp?un=" & myUsers(i).UserName & """>Edit</a></td>" & vbCrLf
                response.Write "<td><a href=""confirmDelete.asp?un=" & myUsers(i).UserName & """>Delete</a></td>" & vbCrLf
            end if
    

            response.Write "</tr>" & vbCrLf
        end if
    next
%>
                        <tr>
                            <td colspan="5" />
                            <td colspan="2">
                                <a href="create.asp">Add New User</a>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </body>
</html>
