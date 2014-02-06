<%
class PartyInviteEmail
    ' attributes
    private mo_Invite

    ' ctor
    public sub Class_Initialize()
        set mo_Invite = new Invitation
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
        set oStream = oFSO.OpenTextFile(Server.MapPath(PathCombine(Settings(SETTING_EMAILTEMPLATE_PATH), Settings(SETTING_EMAILTEMPLATE_FILE_INVITE))))
        body = oStream.ReadAll() & "<div>" & settings(SETTING_EMAILTEMPLATE_INVITE) & "</div>"
        oStream.Close

        GetBody = ReplaceBodyConstants(header & body & footer)
    end function

    private function ReplaceBodyConstants(BodyContent)
        dim oEvent, oEventType

        set oEvent = new ScheduledEvent
        oEvent.Load mo_Invite.EventId

        set oEventType = new EventType
        oEventType.Load oEvent.EventTypeId
    
        BodyContent = NullableReplace(BodyContent, "{{YourName}}", mo_Invite.Name)
        BodyContent = NullableReplace(BodyContent, "{{EventName}}", oEvent.PartyName)
        BodyContent = NullableReplace(BodyContent, "{{EventId}}", EncodeId(oEvent.EventId))
        BodyContent = NullableReplace(BodyContent, "{{Date}}", oEvent.SelectedDate)
        BodyContent = NullableReplace(BodyContent, "{{StartTime}}", oEvent.StartTimeLongFormat)
        BodyContent = NullableReplace(BodyContent, "{{NumberGuests}}", oEvent.NumberOfPatrons)
        BodyContent = NullableReplace(BodyContent, "{{AverageAge}}", )
        BodyContent = NullableReplace(BodyContent, "{{Phone}}", )
        BodyContent = NullableReplace(BodyContent, "{{Subject}}", )
        BodyContent = NullableReplace(BodyContent, "{{GunSize}}", )
        BodyContent = NullableReplace(BodyContent, "{{UserComments}}", )
        BodyContent = NullableReplace(BodyContent, "{{InviteId}}", EncodeId(mo_Invite.InvitationId))
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

        ReplaceBodyConstants = tempBody
    end function

    public sub SendById(nInviteId)
        mo_Invite.Load nInviteId

        if isNull(mo_Invite.EmailAddress) or len(mo_Invite.EmailAddress) = 0 then
            dim failureEmail : set failureEmail = new NoAddressEmail
            failureEmail.Send "Party Invite", mo_Invite.InvitationId
        else
            dim myPrimaryEmail, body
            set myPrimaryEmail = new EmailConnection
        
            myPrimaryEmail.ServerAddress = settings(SETTING_EMAILSERVER)
            myPrimaryEmail.SenderAddress = settings(SETTING_FROMADDRESS)
            myPrimaryEmail.SenderName = settings(SETTING_FROMNAME)
        
            body = GetBody()

            myPrimaryEmail.Send mo_Invite.Name, mo_Invite.EmailAddress, Settings(SETTING_EMAILSUBJECT_INVITE), body
        end if
    end sub

    public sub SendByObject(oInvite)
        if isNull(oInvite.EmailAddress) or len(oInvite.EmailAddress) = 0 then
            dim failureEmail : set failureEmail = new NoAddressEmail
            failureEmail.Send "Party Invite", oInvite.InvitationId
        else
            dim myPrimaryEmail, body
            set myPrimaryEmail = new EmailConnection
            set mo_Invite = oInvite

            myPrimaryEmail.ServerAddress = settings(SETTING_EMAILSERVER)
            myPrimaryEmail.SenderAddress = settings(SETTING_FROMADDRESS)
            myPrimaryEmail.SenderName = settings(SETTING_FROMNAME)

            body = GetBody()

            myPrimaryEmail.Send mo_Invite.Name, mo_Invite.EmailAddress, Settings(SETTING_EMAILSUBJECT_INVITE), body
        end if
    end sub

end class
%>