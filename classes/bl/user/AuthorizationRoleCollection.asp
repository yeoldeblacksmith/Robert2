<%
class AuthorizationRoleCollection
    'attributes 
        private mo_List

    ' ctor
    public sub Class_Initialize()
        set mo_List = new ArrayList
    end sub

    'methods
    public sub Load
        dim myCon
        set myCon = new RoleDataConnection

        dim roles
        roles = myCon.GetAllRoles()

        if UBound(roles) > 0 then
            for i = 0 to UBound(roles, 2)
                dim myRole
                set myRole = new AuthorizationRole

                myRole.RoleId = roles(ROLE_INDEX_ID, i)
                myRole.Description = roles(ROLE_INDEX_DESCRIPTION, i)

                mo_List.Add myRole
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