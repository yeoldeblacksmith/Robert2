<%
class WaiverSearchEmail
    ' attributes
    private mo_Waiver, s_emailTemplate

    ' ctor
    public sub Class_Initialize()
        set mo_Waiver = new Waiver
        s_emailTemplate = settings(SETTING_EMAILTEMPLATE_SEARCHWAIVER)
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
        set oStream = oFSO.OpenTextFile(Server.MapPath(PathCombine(Settings(SETTING_EMAILTEMPLATE_PATH), Settings(SETTING_EMAILTEMPLATE_FILE_SEARCHWAIVER))))
        body = oStream.ReadAll() & "<div>" & s_emailTemplate & "</div>"
        oStream.Close

        GetBody = header & body & footer
    end function
        
    private function GetSubject()
        GetSubject = Settings(SETTING_EMAILSUBJECT_SEARCHWAIVER)
    end function

    private function ReplaceBodyConstants(BodyContent, FName, LName, WaiverLines)
        BodyContent = NullableReplace(BodyContent, "{{WaiverHRefs}}", WaiverLines)
        BodyContent = NullableReplace(BodyContent, "{{YourName}}", Fname & " " & LName)
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
    
    public function GetPreview(subject, body)
        mo_Waiver.LoadById Settings(SETTING_DEFAULT_WAIVERID)

        If len(body) > 0 then
            s_emailTemplate = body
        end if

        GetPreview = ReplaceBodyConstants(GetBody())
    end function

    public sub SendById(sHashId)
        mo_Waiver.Load sHashId

        if isNull(mo_Waiver.EmailAddress) or len(mo_Waiver.EmailAddress) = 0 then
            dim failureEmail : set failureEmail = new NoAddressEmail
            failureEmail.Send "Waiver", sHashId
        else
            dim myEmailCon
            'set myEmailCon = new WaiverEmailConnection
            set myEmailCon = new EmailConnection

            myEmailCon.ServerAddress = settings(SETTING_EMAILSERVER)
            myEmailCon.SenderAddress = settings(SETTING_FROMADDRESS)
            myEmailCon.SenderName = settings(SETTING_FROMNAME)

            myEmailCon.Send mo_Waiver.EmailAddress, mo_Waiver.EmailAddress, GetSubject(), ReplaceBodyConstants(GetBody())
        end if
    end sub

    public sub SendWaiverList(WaiverList, ToAddress)
        dim ToName, WaiverLines
        ToName = UCase(left(WaiverList(0).FirstName,1)) & mid(WaiverList(0).FirstName,2) + " " + UCase(left(WaiverList(0).LastName,1)) & mid(WaiverList(0).LastName,2)

        for i = 0 to WaiverList.Count -1
            WaiverLines = WaiverLines & "<a href=""{{VantoraUrl}}/waiver/display.asp?id=" & WaiverList(i).HashId & """>View online waiver for " & UCase(left(WaiverList(i).FirstName,1)) & mid(WaiverList(i).FirstName,2) & " " & ucase(left(WaiverList(i).LastName,1)) & mid(WaiverList(i).LastName,2) & "</a>.<br />"
        next

        dim myEmailCon
        set myEmailCon = new EmailConnection

        myEmailCon.ServerAddress = settings(SETTING_EMAILSERVER)
        myEmailCon.SenderAddress = settings(SETTING_FROMADDRESS)
        myEmailCon.SenderName = settings(SETTING_FROMNAME)
    
        myEmailCon.Send ToName, ToAddress, GetSubject(),  ReplaceBodyConstants(GetBody(),WaiverList(0).FirstName,WaiverList(0).LastName, WaiverLines)   

    end sub

    public sub SendByHashId(HashId)
        mo_Waiver.Load HashId

        dim ToName, WaiverLine
        ToName = UCase(left(mo_Waiver.FirstName,1)) & mid(mo_Waiver.FirstName,2) + " " + UCase(left(mo_Waiver.LastName,1)) & mid(mo_Waiver.LastName,2)

        WaiverLine = WaiverLine & "<a href=""{{VantoraUrl}}/waiver/display.asp?id=" & mo_Waiver.HashId & """>View online waiver for " & UCase(left(mo_Waiver.FirstName,1)) & mid(mo_Waiver.FirstName,2) & " " & ucase(left(mo_Waiver.LastName,1)) & mid(mo_Waiver.LastName,2) & "</a>.<br />"

        dim myEmailCon
        set myEmailCon = new EmailConnection

        myEmailCon.ServerAddress = settings(SETTING_EMAILSERVER)
        myEmailCon.SenderAddress = settings(SETTING_FROMADDRESS)
        myEmailCon.SenderName = settings(SETTING_FROMNAME)
    
        myEmailCon.Send ToName, mo_Waiver.emailAddress, GetSubject(),  ReplaceBodyConstants(GetBody(),mo_Waiver.FirstName,mo_Waiver.LastName, WaiverLine)   

    end sub

end class
%>