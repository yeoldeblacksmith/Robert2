<%
class User
    ' attributes
    private ms_SiteGuid, ms_UserName, ms_Password, mo_Role, _
            ms_EmailAddress, mb_Enabled, mb_PasswordExpired, ms_PasswordResetHash

    ' ctor
    public sub Class_Initialize()
        set mo_Role = new AuthorizationRole
    end sub

    ' methods
    public function AuthenticateUser(UserName, Password)
        dim tempUser, md5
        set tempUser = new User
        set md5 = new Md5Generator

        tempUser.UserName = UserName
        tempUser.Load
    
        AuthenticateUser = cbool(tempUser.Password = md5.CreateHash(Password & PASSWORD_SALT))
    end function

    public sub ChangePassword()
        dim myCon, md5
        set myCon = new UserDataConnection
        set md5 = new Md5Generator
        
        password = md5.CreateHash(password & PASSWORD_SALT)
        
        myCon.ChangePassword UserName, password
    end sub

    public sub Delete()
        dim myCon
        set myCon = new UserDataConnection

        myCon.DeleteUser UserName
    end sub

    public sub Load()
        dim myCon
        set myCon = new UserDataConnection

        dim results
        results = myCon.GetUser(UserName)

        if ubound(results) > 0 then
            SiteGuid = results(USER_INDEX_SITEGUID, 0)
            UserName = results(USER_INDEX_USERNAME, 0)
            Password = results(USER_INDEX_PASSWORD, 0)
            Role.RoleId = results(USER_INDEX_ROLEID, 0)
            EmailAddress = results(USER_INDEX_EMAILADDRESS, 0)
            Enabled = results(USER_INDEX_ENABLED, 0)
            PasswordExpired = results(USER_INDEX_PASSWORDEXPIRED, 0)
        end if
    end sub

    public sub LoadByEmail(Email)
        dim myCon
        set myCon = new UserDataConnection

        dim results
        results = myCon.GetUserbyEmail(Email)

        if ubound(results) > 0 then
            SiteGuid = results(USER_INDEX_SITEGUID, 0)
            UserName = results(USER_INDEX_USERNAME, 0)
            Password = results(USER_INDEX_PASSWORD, 0)
            Role.RoleId = results(USER_INDEX_ROLEID, 0)
            EmailAddress = results(USER_INDEX_EMAILADDRESS, 0)
            Enabled = results(USER_INDEX_ENABLED, 0)
            PasswordExpired = results(USER_INDEX_PASSWORDEXPIRED, 0)
        end if
    end sub

    public sub LoadByUserHash()
        dim myCon
        set myCon = new UserDataConnection

        dim results
        results = myCon.GetUserByHash(PasswordResetHash)

        if ubound(results) > 0 then
            SiteGuid = results(USER_INDEX_SITEGUID, 0)
            UserName = results(USER_INDEX_USERNAME, 0)
            Password = results(USER_INDEX_PASSWORD, 0)
            Role.RoleId = results(USER_INDEX_ROLEID, 0)
            EmailAddress = results(USER_INDEX_EMAILADDRESS, 0)
            Enabled = results(USER_INDEX_ENABLED, 0)
            PasswordExpired = results(USER_INDEX_PASSWORDEXPIRED, 0)
        end if
    end sub

    public function SaveResetPasswordHash(UserName, SiteGuid)
        dim tempUser, returnValue, md5, myCon, email
        set tempUser = new User
        set email = new PasswordResetHashEmail
        set md5 = new Md5Generator
        set myCon = new UserDataConnection

        returnValue = false

        tempUser.UserName = UserName
        tempUser.Load

        response.write(tempUser.PasswordResetHash)
        if tempUser.SiteGuid = SiteGuid Or tempUser.Role.RoleId = cint(AUTHORIZED_ROLE_GLOBALADMIN) then
            tempUser.PasswordResetHash = md5.CreateHash(UserName)
            myCon.SavePasswordResetHash tempUser.UserName, tempUser.PasswordResetHash, tempUser.SiteGuid
            email.SendByObject tempUser
            returnValue = true
        end if

        SaveResetPasswordHash = cbool(returnValue)
    end function

    public sub Save()
        dim myCon, md5, tempSiteGuid
        set myCon = new UserDataConnection
        set md5 = new Md5Generator

        ' If the user exists preserve their site guid, otherwise use the current sites guid
        tempSiteGuid = MY_SITE_GUID
        if len(SiteGuid) > 0 then
            tempSiteGuid = SiteGuid
        end if

        myCon.SaveUser UserName, Role.RoleId, EmailAddress, Enabled, PasswordExpired, tempSiteGuid
        
        ' Save the password if it exists
        if len(Password) > 0 then
            Password = md5.CreateHash(password & PASSWORD_SALT)
            myCon.ChangePassword UserName, Password
        end if
    end sub

    ' properties
    public property Get SiteGuid
        SiteGuid = ms_SiteGuid
    end property

    public property let SiteGuid(value)
        ms_SiteGuid = value
    end property

    public property Get UserName
        UserName = ms_UserName
    end property

    public property Let UserName(value)
        ms_UserName = value
    end property

    public property Get Password
        Password = ms_Password
    end property

    public property Let Password(value)
        ms_Password = value
    end property

    public property Get Role
        set Role = mo_Role
    end property

    public property let Role(value)
        set mo_Role = value
    end property

    public property Get EmailAddress
        EmailAddress = ms_EmailAddress
    end property

    public property Let EmailAddress(value)
        ms_EmailAddress = value
    end property

    public property Get Enabled
        Enabled = mb_Enabled
    end property

    public property Let Enabled(value)
        mb_Enabled = value
    end property

    public property Get PasswordExpired
        PasswordExpired = mb_PasswordExpired
    end property

    public property Let PasswordExpired(value)
        mb_PasswordExpired = value
    end property

    public property Get PasswordResetHash
        PasswordResetHash = ms_PasswordResetHash
    end property

    public property Let PasswordResetHash(value)
        ms_PasswordResetHash = value
    end property
end class
%>