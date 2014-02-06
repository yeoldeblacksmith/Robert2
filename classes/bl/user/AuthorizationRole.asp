<%
class AuthorizationRole
    ' attributes
    private mn_RoleId, ms_Description

    ' methods
    public sub Load()
        if len(mn_RoleId) = 0 then exit sub

        dim myCon
        set myCon = new RoleDataConnection

        dim results
        results = myCon.GetRole(mn_RoleId)

        if UBound(results) > 0 then
            RoleId = results(ROLE_INDEX_ID, 0)
            Description = results(ROLE_INDEX_DESCRIPTION, 0)
        end if
    end sub

    ' properties
    public property get RoleId
        RoleId = mn_RoleId
    end property

    public property let RoleId(value)
        mn_RoleId = value
    end property

    public property get Description
        Description = ms_Description
    end property

    public property let Description(value)
        ms_Description = value
    end property
end class
%>