<%
class LookupEmail
    ' attributes
    private mo_Waiver, s_emailTemplate

    ' ctor
    public sub Class_Initialize()
        set mo_Waiver = new Waiver
        s_emailTemplate = settings(SETTING_EMAILTEMPLATE_WAIVERLOOKUP)
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
        set oStream = oFSO.OpenTextFile(Server.MapPath(PathCombine(Settings(SETTING_EMAILTEMPLATE_PATH), Settings(SETTING_EMAILTEMPLATE_FILE_WAIVERLOOKUP))))
        body = oStream.ReadAll() & "<div>" & s_emailTemplate & "</div>"
        oStream.Close

        GetBody = ReplaceBodyConstants(header & body & footer)
    end function

    public function GetPreview(subject, body)
        mo_Waiver.LoadById settings(SETTING_DEFAULT_WAIVERID)

        if len(body) > 0 then
            s_emailTemplate = body
        end if

        GetPreview = GetBody()
    end function
        
    private function GetSubject()
        GetSubject = Settings(SETTING_EMAILSUBJECT_WAIVERLOOKUP)
    end function

    private function ReplaceBodyConstants(BodyContent)
        BodyContent = NullableReplace(BodyContent, "{{WaiverId}}", mo_Waiver.HashId)
        BodyContent = NullableReplace(BodyContent, "{{YourName}}", mo_Waiver.FirstName & " " & mo_Waiver.LastName)
        BodyContent = NullableReplace(BodyContent, "{{CompanyName}}", SiteInfo.Name)
        BodyContent = NullableReplace(BodyContent, "{{CompanyAddress}}", SiteInfo.Address)
        BodyContent = NullableReplace(BodyContent, "{{CompanyCity}}", SiteInfo.City)
        BodyContent = NullableReplace(BodyContent, "{{CompanyState}}", SiteInfo.State)
        BodyContent = NullableReplace(BodyContent, "{{CompanyZipCode}}", SiteInfo.ZipCode)
        BodyContent = NullableReplace(BodyContent, "{{CompanyPhone}}", SiteInfo.PhoneNumber)
        BodyContent = NullableReplace(BodyContent, "{{CompanyEmail}}", Settings(SETTING_INFOADDRESS))
        BodyContent = NullableReplace(BodyContent, "{{CompanyUrl}}", SiteInfo.HomeUrl)
        BodyContent = NullableReplace(BodyContent, "{{VantoraUrl}}", SiteInfo.VantoraUrl)
        BodyContent = NullableReplace(BodyContent, "{{DirectoryName}}", SiteInfo.DirectoryName)
        BodyContent = NullableReplace(BodyContent, "{{LogoUrl}}", LOGO_URL)

        ReplaceBodyConstants = BodyContent
    end function

    public sub SendById(sHashId)
        mo_Waiver.Load sHashId
    
        if isNull(mo_Waiver.EmailAddress) or len(mo_Waiver.EmailAddress) = 0 then
            dim failureEmail : set failureEmail = new NoAddressEmail
            failureEmail.Send "Waiver Lookup", sHashId
        else
            dim myEmailCon
            'set myEmailCon = new WaiverEmailConnection
            set myEmailCon = new EmailConnection

            myEmailCon.ServerAddress = settings(SETTING_EMAILSERVER)
            myEmailCon.SenderAddress = settings(SETTING_FROMADDRESS)
            myEmailCon.SenderName = settings(SETTING_FROMNAME)

            myEmailCon.Send mo_Waiver.EmailAddress, mo_Waiver.EmailAddress, GetSubject(), GetBody()
        end if
    end sub
end class
%>