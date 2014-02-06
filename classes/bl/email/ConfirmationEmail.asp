<%
class ConfirmationEmail
    ' attributes
    private mo_Event, s_emailTemplate

    ' ctor
    public sub Class_Initialize()
        set mo_Event = new ScheduledEvent
        s_emailTemplate = settings(SETTING_EMAILTEMPLATE_CONFIRMATION)
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
        set oStream = oFSO.OpenTextFile(Server.MapPath(PathCombine(Settings(SETTING_EMAILTEMPLATE_PATH), Settings(SETTING_EMAILTEMPLATE_FILE_CONFIRMATION))))
        body = oStream.ReadAll() & "<div>" & s_emailTemplate & "</div>"
        oStream.Close

        GetBody = ReplaceBodyConstants(header & body & footer)
    end function

    public function GetPreview(subject, body)
        mo_Event.Load Settings(SETTING_DEFAULT_EVENTID)
        
        if len(body) > 0 then
            s_emailTemplate = body
        end if

        GetPreview = GetBody()
    end function

    private function ReplaceBodyConstants(BodyContent)
        
        'BodyContent = ReplaceIIFs(BodyContent)
        BodyContent = NullableReplace(BodyContent, "{{ModifyEditBlurb}}", ReplaceModifyInfo())
        BodyContent = NullableReplace(BodyContent, "{{YourName}}", mo_Event.ContactName)
        BodyContent = NullableReplace(BodyContent, "{{EventName}}", mo_Event.PartyName)
        BodyContent = NullableReplace(BodyContent, "{{Date}}", mo_Event.SelectedDate)
        BodyContent = NullableReplace(BodyContent, "{{StartTime}}", mo_Event.StartTimeLongFormat)
        BodyContent = NullableReplace(BodyContent, "{{NumberGuests}}", mo_Event.NumberOfPatrons)
        BodyContent = NullableReplace(BodyContent, "{{AverageAge}}", mo_Event.AgeOfPatrons)
        BodyContent = NullableReplace(BodyContent, "{{Phone}}", mo_Event.ContactPhone)
        BodyContent = NullableReplace(BodyContent, "{{Subject}}", Settings(SETTING_EMAILSUBJECT_CONFIRMATION))
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

        dim details : set details = new OrderDetailCollection
        details.LoadByEventId mo_Event.EventId

        if details.TotalAmountOwed = 0 then
            BodyContent = Replace(BodyContent, "{{OrderDetails}}", "")
        else
            BodyContent = Replace(BodyContent, "{{OrderDetails}}", ReplaceOrderDetails(details))
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

    private function ReplaceModifyInfo()
        dim AllowsCancelOrEdit
        AllowsCancelOrEdit = (cbool(Settings(SETTING_REGISTRATION_ALLOWCANCELLATION)) and cint(Settings(SETTING_REGISTRATION_MAXDAYS_CANCEL)) >= 0) or _
                             (cbool(Settings(SETTING_REGISTRATION_ALLOWEDIT)) and cint(Settings(SETTING_REGISTRATION_MAXDAYS_EDIT)) >= 0 )
        dim body
        
        if AllowsCancelOrEdit then
            'if UsesPayments() then
            '    body = "<br/><br/><font size=+1><b>Need To Make A Change??</b></font><br/>" & _
            '           "Call us at {{CompanyPhone}}."
            'else
                body = "<br/><br/><font size=+1><b>Need To Make A Change??</b></font><br/>" & _
                        "You can modify this reservation at any time by following this link:<br/>" & _
                        "<a href=""{{VantoraUrl}}/registration/editregistration.asp?id={{EventId}}"">Modify My Reservation</a>" & _
                        "<br/>Then just log in with your email." & _
                        "<br/><br/>"
            'end if
        else 
            body = "<br/><br/><font size=+1><b>Need To Make A Change??</b></font><br/>" & _
                   "Call us at {{CompanyPhone}}."
        end if

        ReplaceModifyInfo = body
    end function

    private function ReplaceOrderDetails(OrderDetails)
        dim oFSO, oStream, body
        set oFSO = Server.CreateObject("Scripting.FileSystemObject")
        set oStream = oFSO.OpenTextFile(Server.MapPath(PathCombine(Settings(SETTING_EMAILTEMPLATE_PATH), Settings(SETTING_EMAILTEMPLATE_FILE_ORDERDETAILS))))
        body = oStream.ReadAll()
        oStream.Close

        dim rows
        dim adjustments : set adjustments = new ArrayList
        dim originalTotal : originalTotal = 0.00

        for i = 0 to OrderDetails.Count - 1
            if OrderDetails(i).IsAdjustment then
                adjustments.Add OrderDetails(i).ItemNumber
            else
                rows = rows & "<tr style=""vertical-align: top;"">" & vbCrLf
                rows = rows & "<td style=""border-right: 1px solid #aaa; padding: 0px 3px 0px 3px; text-align: left"">" & OrderDetails(i).ItemDescription & "</td>" & vbCrLf
                rows = rows & "<td style=""border-right: 1px solid #aaa; padding: 0px 3px 0px 3px;"">" & OrderDetails(i).Quantity & "</td>" & vbCrLf
                rows = rows & "<td style=""border-right: 1px solid #aaa; padding: 0px 3px 0px 3px;"">" & FormatNumber(OrderDetails(i).Price, 2) & "</td>" & vbCrLf
                rows = rows & "<td style=""padding: 0px 3px 0px 3px;"">" & FormatNumber(OrderDetails(i).ExtendedPrice, 2) & "</td>" & vbCrLf
                rows = rows & "</tr>" & vbCrLf

                originalTotal = originalTotal + OrderDetails(i).ExtendedPrice
            end if
        next

        rows = rows & "<tr>"
		rows = rows & "<td colspan=""3"" style=""border-top: 1px solid #aaa; text-align: left"">Total</td>"
		rows = rows & "<td style=""border-top: 1px solid #aaa;"">" & FormatNumber(originalTotal, 2) & "</td>"
	    rows = rows & "</tr>"

        if adjustments.Count > 0 then
            dim adjustedTotal : adjustedTotal = 0.00
            rows = rows & "<tr><td colspan='4' style='text-align: center'><h3>Modified Order Details</h3></td></tr>"

            for i = 0 to OrderDetails.Count - 1
                dim writableRecord : writableRecord = false

                if OrderDetails(i).IsAdjustment then
                    writableRecord = true
                else
                    dim adjustmentFound : adjustmentFound = false

                    for j = 0 to adjustments.Count - 1
                        if OrderDetails(i).ItemNumber = adjustments(j) then
                            adjustmentFound = true
                            j = adjustments.Count
                        end if
                    next

                    if adjustmentFound = false then writableRecord = true
                end if

                if writableRecord then
                    rows = rows & "<tr style=""vertical-align: top;"">" & vbCrLf
                    rows = rows & "<td style=""border-right: 1px solid #aaa; padding: 0px 3px 0px 3px; text-align: left"">" & OrderDetails(i).ItemDescription & "</td>" & vbCrLf
                    rows = rows & "<td style=""border-right: 1px solid #aaa; padding: 0px 3px 0px 3px;"">" & OrderDetails(i).Quantity & "</td>" & vbCrLf
                    rows = rows & "<td style=""border-right: 1px solid #aaa; padding: 0px 3px 0px 3px;"">" & FormatNumber(OrderDetails(i).Price, 2) & "</td>" & vbCrLf
                    rows = rows & "<td style=""padding: 0px 3px 0px 3px;"">" & FormatNumber(OrderDetails(i).ExtendedPrice, 2) & "</td>" & vbCrLf
                    rows = rows & "</tr>" & vbCrLf

                    adjustedTotal = adjustedTotal + OrderDetails(i).ExtendedPrice
                end if
            next

            rows = rows & "<tr>"
		    rows = rows & "<td colspan=""3"" style=""border-top: 1px solid #aaa; text-align: left"">Total</td>"
		    rows = rows & "<td style=""border-top: 1px solid #aaa;"">" & FormatNumber(adjustedTotal, 2) & "</td>"
	        rows = rows & "</tr>"

            rows = rows & "<tr>"
		    rows = rows & "<td colspan=""3"" style=""text-align: left"">Difference</td>"
		    rows = rows & "<td>" & FormatNumber(originalTotal - adjustedTotal, 2) & "</td>"
	        rows = rows & "</tr>"

            rows = rows & "<tr>"
		    rows = rows & "<td colspan=""4"" style=""text-align: center; color: red"">The difference will be resolved at the field on the day of play</td>"
	        rows = rows & "</tr>"            
        end if

        body = NullableReplace(body, "{{OrderDetails}}", rows)

        ReplaceOrderDetails = body
    end function

    public sub SendById(nEventId)
        mo_Event.Load nEventId
        
        if isNull(mo_Event.ContactEmailAddress) or len(mo_Event.ContactEmailAddress) = 0 then
            dim failureEmail : set failureEmail = new NoAddressEmail
            failureEmail.Send "Event Confirmation", mo_Event.EventId
        else
            dim myPrimaryEmail, mySecondaryEmail, body
            set myPrimaryEmail = new EmailConnection
            set mySecondaryEmail = new EmailConnection

            myPrimaryEmail.ServerAddress = settings(SETTING_EMAILSERVER)
            myPrimaryEmail.SenderAddress = settings(SETTING_FROMADDRESS)
            myPrimaryEmail.SenderName = settings(SETTING_FROMNAME)

            mySecondaryEmail.ServerAddress = settings(SETTING_EMAILSERVER)
            mySecondaryEmail.InfoAddress = settings(SETTING_INFOADDRESS)
            mySecondaryEmail.SenderName = settings(SETTING_FROMNAME)

            body = GetBody()

            myPrimaryEmail.Send mo_Event.ContactName, mo_Event.ContactEmailAddress, Settings(SETTING_EMAILSUBJECT_CONFIRMATION), body
            mySecondaryEmail.SendToReservation mo_Event.ContactName, mo_Event.ContactEmailAddress, Settings(SETTING_EMAILSUBJECT_CONFIRMATION), body
        end if        
    end sub

    public sub SendByObject(oEvent)
        if isNull(oEvent.ContactEmailAddress) or len(oEvent.ContactEmailAddress) = 0 then
            dim failureEmail : set failureEmail = new NoAddressEmail
            failureEmail.Send "Event Confirmation", oEvent.EventId
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
 
            myPrimaryEmail.Send oEvent.ContactName, oEvent.ContactEmailAddress, Settings(SETTING_EMAILSUBJECT_CONFIRMATION), body
            mySecondaryEmail.SendToReservation oEvent.ContactName, oEvent.ContactEmailAddress, Settings(SETTING_EMAILSUBJECT_CONFIRMATION), body
        end if
    end sub
end class
%>