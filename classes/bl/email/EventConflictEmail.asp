<%
class EventConflictEmail
    ' attributes
    private mo_Event

    ' ctor
    public sub Class_Initialize()
        set mo_Event = new ScheduledEvent
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
        set oStream = oFSO.OpenTextFile(Server.MapPath(PathCombine(Settings(SETTING_EMAILTEMPLATE_PATH), Settings(SETTING_EMAILTEMPLATE_FILE_CONFLICT))))
        body = oStream.ReadAll() & "<div>" & settings(SETTING_EMAILTEMPLATE_CONFLICT) & "</div>"
        oStream.Close

        GetBody = ReplaceBodyConstants(header & body & footer)
    end function

    private function ReplaceBodyConstants(BodyContent)
        BodyContent = NullableReplace(BodyContent, "{{YourName}}", mo_Event.ContactName)
        BodyContent = NullableReplace(BodyContent, "{{EventName}}", mo_Event.PartyName)
        BodyContent = NullableReplace(BodyContent, "{{Date}}", mo_Event.SelectedDate)
        BodyContent = NullableReplace(BodyContent, "{{StartTime}}", mo_Event.StartTimeLongFormat)
        BodyContent = NullableReplace(BodyContent, "{{NumberGuests}}", mo_Event.NumberOfPatrons)
        BodyContent = NullableReplace(BodyContent, "{{AverageAge}}", mo_Event.AgeOfPatrons)
        BodyContent = NullableReplace(BodyContent, "{{Phone}}", mo_Event.ContactPhone)
        BodyContent = NullableReplace(BodyContent, "{{Subject}}", Settings(SETTING_EMAILSUBJECT_CONFLICT))
        BodyContent = NullableReplace(BodyContent, "{{UserComments}}", mo_Event.UserComments)
        BodyContent = NullableReplace(BodyContent, "{{EventId}}", EncodeId(mo_Event.EventId))
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

    public sub SendById(nEventId)
        mo_Event.Load nEventId
    
        if isNull(mo_Event.ContactEmailAddress) or len(mo_Event.ContactEmailAddress) = 0 then
            dim failureEmail : set failureEmail = new NoAddressEmail
            failureEmail.Send "Event Conflict", mo_Event.EventId
        else
            dim myPrimaryEmail, mySecondaryEmail
            set myPrimaryEmail = new EmailConnection
            set mySecondaryEmail = new EmailConnection

            myPrimaryEmail.ServerAddress = settings(SETTING_EMAILSERVER)
            myPrimaryEmail.SenderAddress = settings(SETTING_FROMADDRESS)
            myPrimaryEmail.SenderName = settings(SETTING_FROMNAME)

            mySecondaryEmail.ServerAddress = settings(SETTING_EMAILSERVER)
            mySecondaryEmail.InfoAddress = settings(SETTING_INFOADDRESS)
            mySecondaryEmail.SenderName = settings(SETTING_FROMNAME)

            dim body : body = GetBody()

            myPrimaryEmail.Send mo_Event.ContactName, mo_Event.ContactEmailAddress, Settings(SETTING_EMAILSUBJECT_CONFLICT), body
            mySecondaryEmail.SendToReservation mo_Event.ContactName, mo_Event.ContactEmailAddress, Settings(SETTING_EMAILSUBJECT_CONFLICT), body
        end if
    end sub

    public sub SendByObject(oEvent)
        if isNull(oEvent.ContactEmailAddress) or len(oEvent.ContactEmailAddress) = 0 then
            dim failureEmail : set failureEmail = new NoAddressEmail
            failureEmail.Send "Event Conflict", oEvent.EventId
        else
            dim myPrimaryEmail, mySecondaryEmail, body
            set myPrimaryEmail = new EmailConnection
            set mySecondaryEmail = new EmailConnection
            set mo_Event = oEvent

            myPrimaryEmail.ServerAddress = settings(SETTING_EMAILSERVER)
            myPrimaryEmail.SenderAddress = settings(SETTING_FROMADDRESS)
            myPrimaryEmail.SenderName = settings(SETTING_FROMNAME)

            mySecondaryEmail.ServerAddress = settings(SETTING_EMAILSERVER)
            mySecondaryEmail.InfoAddress = settings(SETTING_INFOADDRESS)
            mySecondaryEmail.SenderName = settings(SETTING_FROMNAME)


            body = GetBody()

            myPrimaryEmail.Send oEvent.ContactName, oEvent.ContactEmailAddress, Settings(SETTING_EMAILSUBJECT_CONFLICT), body
            mySecondaryEmail.SendToReservation oEvent.ContactName, oEvent.ContactEmailAddress, Settings(SETTING_EMAILSUBJECT_CONFLICT), body
        end if
    end sub
end class
%>