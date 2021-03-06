﻿<%
class ReminderEmail
    ' attributes
    private mo_Event, s_emailTemplate

    ' ctor
    public sub Class_Initialize()
        set mo_Event = new ScheduledEvent
        s_emailTemplate = settings(SETTING_EMAILTEMPLATE_REMINDER)
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
        set oStream = oFSO.OpenTextFile(Server.MapPath(PathCombine(Settings(SETTING_EMAILTEMPLATE_PATH), Settings(SETTING_EMAILTEMPLATE_FILE_REMINDER))))
        body = oStream.ReadAll() & "<div>" & s_emailTemplate & "</div>"
        oStream.Close

        GetBody = ReplaceBodyConstants(header & body & footer)
    end function

    public function GetPreview(subject, body)
        mo_Event.Load settings(SETTING_DEFAULT_EVENTID)

       if len(body) > 0 then
          s_emailTemplate = body
        end if

        GetPreview = GetBody()
    end function
        
    private function GetSubject()
        GetSubject = Settings(SETTING_EMAILSUBJECT_REMINDER) & " on " & FormatDateTime(mo_Event.SelectedDate, vbLongDate)
    end function

    private function ReplaceBodyConstants(BodyContent)
        BodyContent = NullableReplace(BodyContent, "{{YourName}}", mo_Event.ContactName)
        BodyContent = NullableReplace(BodyContent, "{{EventName}}", mo_Event.PartyName)
        BodyContent = NullableReplace(BodyContent, "{{Date}}", mo_Event.SelectedDate)
        BodyContent = NullableReplace(BodyContent, "{{StartTime}}", mo_Event.StartTimeLongFormat)
        BodyContent = NullableReplace(BodyContent, "{{NumberGuests}}", mo_Event.NumberOfPatrons)
        BodyContent = NullableReplace(BodyContent, "{{AverageAge}}", mo_Event.AgeOfPatrons)
        BodyContent = NullableReplace(BodyContent, "{{Phone}}", mo_Event.ContactPhone)
        BodyContent = NullableReplace(BodyContent, "{{Subject}}", APP_SETTING_CONFIRMATIONSUBJECT)
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

        mo_Event.LoadValues

        if mo_Event.CustomFieldValues.Count = 0 then
            BodyContent = Replace(BodyContent, "{{CustomFields}}", "")
        else
            dim customFieldBody, values
            customFieldBody = ""
            set values = mo_Event.CustomFieldValues

            for i = 0 to values.Count - 1
                dim myValue
                set myValue = values(i)
    
                dim field
                set field = new RegistrationCustomField
                field.RegistrationCustomFieldId = myValue.RegistrationCustomFieldId
                field.Load

                select case field.CustomFieldDataType.CustomFieldDataTypeId
                    case FIELDTYPE_LONGTEXT
                        customFieldBody = customFieldBody & ReplaceCustomFieldsLong(field.Name, myValue.Value)
                    case FIELDTYPE_OPTIONS
                        dim custOption
                        set custOption = new RegistrationCustomOption
                        custOption.RegistrationCustomFieldId = myValue.RegistrationCustomFieldId
                        custOption.Sequence = myValue.Value
                        custOption.Load

                        customFieldBody = customFieldBody & ReplaceCustomFieldsShort(field.Name, custOption.Text)
                    case FIELDTYPE_SHORTTEXT
                        customFieldBody = customFieldBody & ReplaceCustomFieldsShort(field.Name, myValue.Value)
                    case FIELDTYPE_YESNO
                        'if cbool(myValue.Value) then myValue.Value = "Yes" else myValue.Value = "No"
                        customFieldBody = customFieldBody & ReplaceCustomFieldsShort(field.Name, iif(cbool(myValue.Value), "Yes", "No"))
                end select
            next

            BodyContent = Replace(BodyContent, "{{CustomFields}}", customFieldBody)
        end if

        dim waivers, playerSignatures, waiverDictionary
        set waivers = new WaiverCollection
        set waiverDictionary = Server.CreateObject("Scripting.Dictionary")
        waivers.LoadByEventId mo_Event.EventId, 1, 1

        for i = 0 to waivers.Count - 1
            if waiverDictionary.Exists(waivers(i).FirstName + waivers(i).LastName + waivers(i).PlayerDOB) = False then 
                waiverDictionary.Add waivers(i).FirstName + waivers(i).LastName + waivers(i).PlayerDOB, waivers(i).LastName + ", " + waivers(i).FirstName + "<br/>"
            end if
        next

        if waivers.Count = 0 then
            BodyContent = Replace(BodyContent, "{{SignatureList}}", "<b><i>No one has completed a waiver online yet.</i></b><br/><br/>")
        else
            for each key in waiverDictionary.Keys
                playerSignatures = playerSignatures + waiverDictionary(key)
            next

            BodyContent = Replace(BodyContent, "{{SignatureList}}", ReplaceWaiverSignatures(playerSignatures))
        end if

        ReplaceBodyConstants = BodyContent
    end function

    private function ReplaceCustomFieldsLong(FieldName, FieldValue)
        dim oFSO, oStream, body
        set oFSO = Server.CreateObject("Scripting.FileSystemObject")
        set oStream = oFSO.OpenTextFile(Server.MapPath(PathCombine(Settings(SETTING_EMAILTEMPLATE_PATH), Settings(SETTING_EMAILTEMPLATE_FILE_CUSTOMFIELDLONG))))
        body = oStream.ReadAll()
        oStream.Close

        body = NullableReplace(body, "{{FieldName}}", FieldName)
        body = NullableReplace(body, "{{FieldValue}}", FieldValue)

        ReplaceCustomFieldsLong = body
    end function

    private function ReplaceCustomFieldsShort(FieldName, FieldValue)
        dim oFSO, oStream, body
        set oFSO = Server.CreateObject("Scripting.FileSystemObject")
        set oStream = oFSO.OpenTextFile(Server.MapPath(PathCombine(Settings(SETTING_EMAILTEMPLATE_PATH), Settings(SETTING_EMAILTEMPLATE_FILE_CUSTOMFIELDSHORT))))
        body = oStream.ReadAll()
        oStream.Close

        body = NullableReplace(body, "{{FieldName}}", FieldName)
        body = NullableReplace(body, "{{FieldValue}}", FieldValue)

        ReplaceCustomFieldsShort = body
    end function

    private function ReplaceWaiverSignatures(ListOfNames)
        dim oFSO, oStream, body
        set oFSO = Server.CreateObject("Scripting.FileSystemObject")

        set oStream = oFSO.OpenTextFile(Server.MapPath(PathCombine(Settings(SETTING_EMAILTEMPLATE_PATH), Settings(SETTING_EMAILTEMPLATE_FILE_SIGNATURES))))

        body = oStream.ReadAll()
        oStream.Close

        body = NullableReplace(body, "{{Signatures}}", ListOfNames)

        ReplaceWaiverSignatures = body
    end function

    public sub SendById(nEventId)
        mo_Event.Load nEventId

        if isNull(mo_Event.ContactEmailAddress) or len(mo_Event.ContactEmailAddress) = 0 then
            dim failureEmail : set failureEmail = new NoAddressEmail
            failureEmail.Send "Event Reminder", mo_Event.EventId
        else
            dim myEmailCon
            set myEmailCon = new EmailConnection

            myEmailCon.ServerAddress = settings(SETTING_EMAILSERVER)
            myEmailCon.SenderAddress = settings(SETTING_FROMADDRESS)
            myEmailCon.SenderName = settings(SETTING_FROMNAME)

            myEmailCon.Send mo_Event.ContactName, mo_Event.ContactEmailAddress, GetSubject(), GetBody()
        end if
    end sub

    public sub SendByObject(oEvent)
        if isNull(oEvent.ContactEmailAddress) or len(oEvent.ContactEmailAddress) = 0 then
            dim failureEmail : set failureEmail = new NoAddressEmail
            failureEmail.Send "Event Reminder", oEvent.EventId
        else
            dim myEmailCon
            set myEmailCon = new EmailConnection
            set mo_Event = oEvent

            myEmailCon.ServerAddress = settings(SETTING_EMAILSERVER)
            myEmailCon.SenderAddress = settings(SETTING_FROMADDRESS)
            myEmailCon.SenderName = settings(SETTING_FROMNAME)

            myEmailCon.Send oEvent.ContactName, oEvent.ContactEmailAddress, GetSubject(), GetBody()
        end if
    end sub
end class
%>