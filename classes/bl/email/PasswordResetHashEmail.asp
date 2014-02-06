<%
class PasswordResetHashEmail
    ' attributes
    private mo_User
    private ms_NewPassword

    ' ctor
    public sub Class_Initialize()
        set mo_User = new User
    end sub

    ' methods
    private function GetBody()
        dim body, header, footer, oFSO, oStream
        set oFSO = Server.CreateObject("Scripting.FileSystemObject")

        'set oStream = oFSO.OpenTextFile(Server.MapPath(EMAILTEMPLATE_PATH_HEADER))
        set oStream = oFSO.OpenTextFile(Server.MapPath(PathCombine(Settings(SETTING_EMAILTEMPLATE_PATH), Settings(SETTING_EMAILTEMPLATE_FILE_HEADER))))
        header = oStream.ReadAll()
        oStream.Close

        'set oStream = oFSO.OpenTextFile(Server.MapPath(EMAILTEMPLATE_PATH_FOOTER))
        set oStream = oFSO.OpenTextFile(Server.MapPath(PathCombine(Settings(SETTING_EMAILTEMPLATE_PATH), Settings(SETTING_EMAILTEMPLATE_FILE_FOOTER))))
        footer = oStream.ReadAll()
        oStream.Close

        'set oStream = oFSO.OpenTextFile(Server.MapPath(EMAILTEMPLATE_PATH_BODY_CANCEL))
        set oStream = oFSO.OpenTextFile(Server.MapPath(PathCombine(Settings(SETTING_EMAILTEMPLATE_PATH), Settings(SETTING_EMAILTEMPLATE_FILE_PASSWORDRECOVERY))))
        body = oStream.ReadAll() & "<div>" & settings(SETTING_EMAILTEMPLATE_PASSWORDRECOVERY) & "</div>"
        oStream.Close

        GetBody = ReplaceBodyConstants(header & body & footer)
    end function

    private function ReplaceBodyConstants(BodyContent)
        dim tempBody

        tempBody = NullableReplace(BodyContent, "{{YourName}}", mo_User.UserName)
        tempBody = NullableReplace(tempBody, "{{NewPassword}}", ms_NewPassword)
        tempBody = NullableReplace(tempBody, "{{CompanyName}}", SiteInfo.Name)
        tempBody = NullableReplace(tempBody, "{{CompanyAddress}}", SiteInfo.Address)
        tempBody = NullableReplace(tempBody, "{{CompanyCity}}", SiteInfo.City)
        tempBody = NullableReplace(tempBody, "{{CompanyState}}", SiteInfo.State)
        tempBody = NullableReplace(tempBody, "{{CompanyZipCode}}", SiteInfo.ZipCode)
        tempBody = NullableReplace(tempBody, "{{CompanyPhone}}", SiteInfo.PhoneNumber)
        tempBody = NullableReplace(tempBody, "{{CompanyEmail}}", Settings(SETTING_INFOADDRESS))
        tempBody = NullableReplace(tempBody, "{{CompanyUrl}}", SiteInfo.HomeUrl)
        tempBody = NullableReplace(tempBody, "{{ResetPasswordUrl}}", PathCombine(SiteInfo.VantoraUrl, "admin/user/resetpassword.asp?u=" & mo_User.PasswordResetHash))
        tempBody = NullableReplace(tempBody, "{{DirectoryName}}", SiteInfo.DirectoryName)
        tempBody = NullableReplace(tempBody, "{{LogoUrl}}", LOGO_URL)


        ReplaceBodyConstants = tempBody
    end function

    public sub SendByUserName(UserName, Password)
        dim myInvite, myPrimaryEmail, body
        set myPrimaryEmail = new EmailConnection
        
        myPrimaryEmail.ServerAddress = settings(SETTING_EMAILSERVER)
        myPrimaryEmail.SenderAddress = settings(SETTING_FROMADDRESS)
        myPrimaryEmail.SenderName = settings(SETTING_FROMNAME)

        mo_Invite.Load nInviteId
        
        body = GetBody()

        myPrimaryEmail.Send mo_Invite.Name, mo_Invite.EmailAddress, Settings(SETTING_EMAILSUBJECT_INVITE), body
    end sub

    public sub SendByObject(oUser)
        dim myPrimaryEmail, body
        set myPrimaryEmail = new EmailConnection
        set mo_User = oUser

        myPrimaryEmail.ServerAddress = settings(SETTING_EMAILSERVER)
        myPrimaryEmail.SenderAddress = settings(SETTING_FROMADDRESS)
        myPrimaryEmail.SenderName = settings(SETTING_FROMNAME)

        body = GetBody()

        myPrimaryEmail.Send oUser.UserName, oUser.EmailAddress, Settings(SETTING_EMAILSUBJECT_PASSWORDRECOVERY), body
    end sub

end class
%>