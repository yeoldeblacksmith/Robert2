<%
class UserCollection
    'attributes 
    private mo_List

    ' ctor
    public sub Class_Initialize()
        set mo_List = new ArrayList
    end sub

    'methods
    public sub LoadForSite
        dim myCon
        set myCon = new UserDataConnection

        dim users
        users = myCon.GetAllUsersForSite()

        if UBound(users) > 0 then
            for i = 0 to UBound(users, 2)
                dim myUser
                set myUser = new User

                myUser.SiteGuid = users(USER_INDEX_SITEGUID, i)
                myUser.UserName = users(USER_INDEX_USERNAME, i)
                myUser.Password = users(USER_INDEX_PASSWORD, i)
                myUser.Role.RoleId = users(USER_INDEX_ROLEID, i)
                myUser.EmailAddress = users(USER_INDEX_EMAILADDRESS, i)
                myUser.Enabled = users(USER_INDEX_ENABLED, i)
                myUser.PasswordExpired = users(USER_INDEX_PASSWORDEXPIRED, i)

                mo_List.Add myUser
            next
        end if
    end sub

    'properties
    public property get Count
        Count = mo_List.Count
    end property

    Public Default Property Get Item(index)
        set Item = mo_List(index)
    end property
end class
%>