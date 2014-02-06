<!--#include file="../classes/IncludeList.asp"-->
<% 
    CONST AJAXACTION_EVENT_BUILDDAILYTABLE = "1"
    CONST AJAXACTION_EVENT_BUILDPARTYLIST = "8"
    CONST AJAXACTION_EVENT_BUILDREMAININGTABLE = "2"
    CONST AJAXACTION_EVENT_BUILDREPORTARRAY = "7"
    CONST AJAXACTION_EVENT_CANCELEVENT = "4"
    CONST AJAXACTION_EVENT_CHECKEXISTS = "10"
    CONST AJAXACTION_EVENT_CHECKOUTPARTY = "9"
    CONST AJAXACTION_EVENT_MARKASPAID = "3"
    CONST AJAXACTION_EVENT_SHOWEDITFORM = "6"
    CONST AJAXACTION_EVENT_VALIDATEEMAIL = "5"
    const KEY_PAYMENTSTATUS_ERROR = "Error"
    const KEY_PAYMENTSTATUS_PAID = "Paid"
    const KEY_PAYMENTSTATUS_UNPAID = "Unpaid"

    '*****************************************************************
    ' check the querystring to determine what action to take
    '*****************************************************************
    select case Request.QueryString(QUERYSTRING_VAR_ACTION)
        case AJAXACTION_EVENT_BUILDDAILYTABLE
            WriteEventTableForDay Request.QueryString(QUERYSTRING_VAR_SELECTEDDATE)
        case AJAXACTION_EVENT_BUILDPARTYLIST
            ' used in the daily stacked report
            BuildPartyList Request.QueryString(QUERYSTRING_VAR_SELECTEDDATE)
        case AJAXACTION_EVENT_BUILDREMAININGTABLE
            WriteEventTableRemaining
        case AJAXACTION_EVENT_BUILDREPORTARRAY
            ' used for the daily stacked report
            BuildDailyStackedReportDataArray Request.QueryString(QUERYSTRING_VAR_SELECTEDDATE)
        case AJAXACTION_EVENT_CANCELEVENT
            dim cancelEventId
            cancelEventId = Request.QueryString(QUERYSTRING_VAR_ID)
        
            if len(Request.QueryString(QUERYSTRING_VAR_DECODE)) > 0 then
                if request.QueryString(QUERYSTRING_VAR_DECODE) = "true" then
                    cancelEventId = DecodeId(Request.QueryString(QUERYSTRING_VAR_ID))
                end if
            end if

            CancelExistingEvent cancelEventId, Request.QueryString(QUERYSTRING_VAR_STATUS), Request.QueryString(QUERYSTRING_VAR_RESENDEMAIL)
        case AJAXACTION_EVENT_CHECKEXISTS
            CheckEventExistsByEmail Request.QueryString(QUERYSTRING_VAR_EMAIL)
        case AJAXACTION_EVENT_CHECKOUTPARTY
            CheckOutEvent Request.QueryString(QUERYSTRING_VAR_ID), request.QueryString(QUERYSTRING_VAR_STARTTIME)
        case AJAXACTION_EVENT_MARKASPAID
            MarkAsPaid Request.QueryString(QUERYSTRING_VAR_ID)
        case AJAXACTION_EVENT_SHOWEDITFORM
            if VerifyEmail(Request.QueryString(QUERYSTRING_VAR_ID), Request.QueryString(QUERYSTRING_VAR_EMAIL)) then
                ShowRegistrationForm Request.QueryString(QUERYSTRING_VAR_ID)
            else
                ShowRegistrationValidationForm Request.QueryString(QUERYSTRING_VAR_EMAIL), false
            end if
        case AJAXACTION_EVENT_VALIDATEEMAIL
            ShowRegistrationValidationForm "", true
    end select

    '**********************************************************
    ' build the custom fields for the reigstration form
    '**********************************************************
    sub BuildCustomFields(EventId)
        dim fields
        set fields = new RegistrationCustomFieldCollection

        fields.LoadAll

        for i = 0 to fields.Count - 1
            response.Write "<tr>" & vbCrLf
    
            select case fields(i).CustomFieldDataType.CustomFieldDataTypeId
                case FIELDTYPE_SHORTTEXT
                    BuildCustomControlTextField fields(i), EventId
                case FIELDTYPE_LONGTEXT
                    BuildCustomControlTextArea  fields(i), EventId
                case FIELDTYPE_YESNO
                    BuildCustomControlCheckBox fields(i), EventId
                case FIELDTYPE_OPTIONS
                    BuildCustomControlDropDownList fields(i), EventId
            end select
                
            response.Write "</tr>" & vbCrLf
        next
    end sub

    sub BuildCustomControlCheckBox(CurrentField, EventId)
        dim currentValue : currentValue = CurrentField.GetFormValue(EventId)
        dim fieldName : fieldName = "CustomField" & CurrentField.RegistrationCustomFieldId
        dim hasAddtionalInfo : hasAddtionalInfo = cbool(len(CurrentField.AdditionalInformation) > 0)
        dim hasNotes : hasNotes = cbool(len(CurrentField.Notes) > 0)
        dim isRequired : isRequired = CurrentField.Required
        dim validatorName : validatorName = "val" & fieldName

        with response
            .Write "<td style=""text-align: right"">" & CurrentField.Name & ":</td>" & vbCrLf
            .Write "<td>" & vbCrLf
            .Write "<input type=""checkbox"" name=""" & fieldName & """ id=""" & fieldName & """ " & iif(cbool(currentValue), "checked=""checked""", "") & " "

            .Write "class="""
            if hasNotes then .Write " tooltip"
            if isRequired then .Write " requiredCustom"
            .Write """ "

            if hasNotes then .write "data-powertip=""" & nullablereplace(CurrentField.Notes, """", "&quot;") & """ "
            if isRequired then .Write "data-reqValidatorId=""" & validatorName & """"

            .Write "/><br/>" & vbCrLF

            if isRequired then .Write "<span id=""" & validatorName & """ style=""color: red; display: none"">Required</span>" & vbCrLf

            .Write "</td>" & vbCrLf

            if isRequired then .Write "<td><span style=""color: red"">&nbsp; *</span></td>" & vbCrLf else .Write "<td/>" & vbCrLf
            if hasAddtionalInfo then
                .Write "<td>" & vbCrLf
                .Write "<img src=""../content/images/info.gif"" class=""mouseover"" alt="""" title=""additional information"" border=""0"" data-info=""" & _
                        NullableReplace(currentField.AdditionalInformation, """", "&quot;") & """ onclick=""showAdditionalInfo(this)"" />" & vbCrLf
                .Write "</td>" & vbCrLf
            else
                .Write "<td />" & vbCrLf
            end if
        end with
    end sub

    sub BuildCustomControlDropDownList(CurrentField, EventId)
        dim currentValue : currentValue = CurrentField.GetFormValue(EventId)
        dim fieldName : fieldName = "CustomField" & CurrentField.RegistrationCustomFieldId
        dim hasAddtionalInfo : hasAddtionalInfo = cbool(len(CurrentField.AdditionalInformation) > 0)
        dim hasNotes : hasNotes = cbool(len(CurrentField.Notes) > 0)
        dim isRequired : isRequired = CurrentField.Required
        dim validatorName : validatorName = "val" & fieldName

        with response
            .Write "<td style=""text-align: right"">" & CurrentField.Name & ":</td>" & vbCrLf
            .Write "<td>" & vbCrLf
            .Write "<select name=""" & fieldName & """ id=""" & fieldName & """  "

            .Write "class="""
            if hasNotes then .Write " tooltip"
            if isRequired then .Write " requiredCustom"
            .Write """ "

            if hasNotes then .write "data-powertip=""" & nullablereplace(CurrentField.Notes, """", "&quot;") & """ "
            if isRequired then .Write "data-reqValidatorId=""" & validatorName & """"
            if CurrentField.HasPayment then .Write "data-feetype=""" & CurrentField.PaymentTypeId & """ data-feefieldid=""" & CurrentField.RegistrationCustomFieldId & """ onchange=""calculateTotals()"""

            .Write ">" & vbCrLf

            .Write "<option></option>" & vbCrLf

            CurrentField.LoadOptions
            dim options
            set options = CurrentField.Options

            for j = 0 to options.Count - 1
                if CurrentField.HasPayment then
                '    dim optionText : optionText = options(j).Text & " (" & GetPaymentText(CurrentField.PaymentTypeId, options(j).Value) & ")"
                    .Write "<option value=""" & options(j).Sequence & """ data-value=""" & options(j).Value & """ data-text=""" & options(j).Text & """ " & iif(CStr(options(j).Sequence) = currentValue, "selected=""selected""", "") & ">" & options(j).Text & "</option>" & vbCrLf
                '    .Write "<option value=""" & options(j).Sequence & """ data-value=""" & options(j).Value & """ data-text=""" & options(j).Text & """>" & options(j).Text & "</option>" & vbCrLf
                else
                    .Write "<option value=""" & options(j).Sequence & """ " & iif(CStr(options(j).Sequence) = currentValue, "selected=""selected""", "") & ">" & options(j).Text & "</option>" & vbCrLf
                end if
            next

            .Write "</select>" & vbCrLF

            if isRequired then .Write "<span id=""" & validatorName & """ style=""color: red; display: none"">Required</span>" & vbCrLf

            .Write "</td>" & vbCrLf
            if isRequired then 
                .Write "<td style=""color: Red""> &nbsp; *</td>"  & vbCrLf
            else
                .Write "<td/>" & vbCrLf
            end if
            if hasAddtionalInfo then
                .Write "<td>" & vbCrLf
                .Write "<img src=""../content/images/info.gif"" class=""mouseover"" alt="""" title=""additional information"" border=""0"" data-info=""" & _
                        NullableReplace(CurrentField.AdditionalInformation, """", "&quot;") & """ onclick=""showAdditionalInfo(this)"" />" & vbCrLf
                .Write "</td>" & vbCrLf
            else
                .Write "<td/>" & vbCrLf
            end if
        end with
    end sub 

    sub BuildCustomControlTextArea(CurrentField, EventId)
        dim currentValue : currentValue = CurrentField.GetFormValue(EventId)
        dim fieldName : fieldName = "CustomField" & CurrentField.RegistrationCustomFieldId
        dim hasAddtionalInfo : hasAddtionalInfo = cbool(len(CurrentField.AdditionalInformation) > 0)
        dim hasNotes : hasNotes = cbool(len(CurrentField.Notes) > 0)
        dim isRequired : isRequired = CurrentField.Required
        dim validatorName : validatorName = "val" & fieldName

        with response
            .Write "<td style=""text-align: right"">" & CurrentField.Name & ":</td>" & vbCrLf
            .Write "<td/>" & vbCrLf
            if isRequired then 
                .Write "<td style=""color: Red""> &nbsp; *</td>"  & vbCrLf
                if hasAddtionalInfo then
                    .Write "<td>" & vbCrLf
                    .Write "<img src=""../content/images/info.gif"" class=""mouseover"" alt="""" title=""additional information"" border=""0"" data-info=""" & _
                            NullableReplace(CurrentField.AdditionalInformation, """", "&quot;") & """ onclick=""showAdditionalInfo(this)"" />" & vbCrLf
                    .Write "</td>" & vbCrLf
                else
                    .Write "<td/>" & vbCrLf
                end if
            else
                if hasAddtionalInfo then
                    .Write "<td/>"
                    .Write "<td>" & vbCrLf
                    .Write "<img src=""../content/images/info.gif"" class=""mouseover"" alt="""" title=""additional information"" border=""0"" data-info=""" & _
                            NullableReplace(CurrentField.AdditionalInformation, """", "&quot;") & """ onclick=""showAdditionalInfo(this)"" />" & vbCrLf
                    .Write "</td>" & vbCrLf
                else
                    .Write "<td colspan=""3""/>" & vbCrLf
                end if
            end if
            .Write "</tr>" & vbCrLf
            .Write "<tr>" & vbCrLf
            .Write "<td colspan=""4"">" & vbCrLf

            .Write "<textarea id=""" & FieldName & """ name=""" & fieldName & """ style=""resize: none; width: 98%"" "

            .Write "class="""
            if hasNotes then .Write " tooltip"
            if isRequired then .Write " requiredCustom"
            .Write """ "

            if hasNotes then .write "data-powertip=""" & nullablereplace(CurrentField.Notes, """", "&quot;") & """ "
            if isRequired then .Write "data-reqValidatorId=""" & validatorName & """"

            .Write ">" & currentValue & "</textarea>" & vbCrLf
            
            if isRequired then .Write "<span id=""" & validatorName & """ style=""color: red; display: none"">Required</span>" & vbCrLf

            .Write "</td>" & vbCrLf
        end with
    end sub

    sub BuildCustomControlTextField(CurrentField, EventId)
        dim currentValue : currentValue = CurrentField.GetFormValue(EventId)
        dim fieldName : fieldName = "CustomField" & CurrentField.RegistrationCustomFieldId
        dim hasAddtionalInfo : hasAddtionalInfo = cbool(len(CurrentField.AdditionalInformation) > 0)
        dim hasNotes : hasNotes = cbool(len(CurrentField.Notes) > 0)
        dim isRequired : isRequired = CurrentField.Required
        dim validatorName : validatorName = "val" & fieldName

        with response
            .Write "<td style=""text-align: right"">" & CurrentField.Name & ":</td>" & vbCrLf
            .Write "<td>" & vbCrLf
            .Write "<input type=""text"" name=""" & fieldName & """ id=""" & fieldName & """ "

            .Write "class="""
            if hasNotes then .Write " tooltip"
            if isRequired then .Write " requiredCustom"
            .Write """ "

            if hasNotes then .write "data-powertip=""" & nullablereplace(CurrentField.Notes, """", "&quot;") & """ "
            if isRequired then .Write "data-reqValidatorId=""" & validatorName & """"

            .Write " value=""" & currentValue & """/>" & vbCrLF

            if isRequired then .Write "<span id=""" & validatorName & """ style=""color: red; display: none"">Required</span>" & vbCrLf

            .Write "</td>" & vbCrLf
            if isRequired then 
                .Write "<td style=""color: Red""> &nbsp; *</td>"  & vbCrLf
            else
                .Write "<td/>" & vbCrLf
            end if
            if hasAddtionalInfo then
                .Write "<td>" & vbCrLf
                .Write "<img src=""../content/images/info.gif"" class=""mouseover"" alt="""" title=""additional information"" border=""0"" data-info=""" & _
                        NullableReplace(CurrentField.AdditionalInformation, """", "&quot;") & """ onclick=""showAdditionalInfo(this)"" />" & vbCrLf
                .Write "</td>" & vbCrLf
            else
                .Write "<td/>" & vbCrLf
            end if
        end with
    end sub

    '**********************************************************
    ' build the array that is passed to the google data table
    ' used for the stacked bar chart
    '**********************************************************
    sub BuildDailyStackedReportDataArray(SelectedDate)
        dim myEvents, reportJSON
        set myEvents = new ScheduledEventCollection
        myEvents.LoadByDateWithoutCheckout SelectedDate

        reportJSON = BuildDailyStackedReportHeaderColumns(myEvents)
        reportJSON = reportJSON & BuildDailyStackedReportRows(SelectedDate, myEvents)

        response.Write reportJSON
    end sub

    '**********************************************************
    ' build the first line of the array that defines the data
    ' for the daily stacked report
    '**********************************************************
    private function BuildDailyStackedReportHeaderColumns(SelectedEvents)
        dim columnString

        columnString = "{""cols"":[{""id"":"""",""label"":""Time"",""pattern"":"""",""type"":""string""},"

        for i = 0 to SelectedEvents.Count - 1
            if SelectedEvents(i).IsPaid then
                columnString = columnString & "{""id"":"""",""label"":""" 
                if len(SelectedEvents(i).PartyName) > 0 then
                    columnString = columnString &  Replace(SelectedEvents(i).PartyName, """", "'")
                else
                    columnString = columnString & SelectedEvents(i).ContactName
                end if

                columnString = columnString & """,""pattern"":"""",""type"":""number""}"

                if i < SelectedEvents.Count - 1 then
                    columnString = columnString &  ","
                end if
            end if
        next

        columnString = columnString & "],"

        BuildDailyStackedReportHeaderColumns = columnString
    end function

    '**********************************************************
    ' build the detail lines of the array for the daily
    ' stacked report
    '**********************************************************
    private function BuildDailyStackedReportRows(SelectedDate, SelectedEvents)
        dim rowString, myDate, currentTime
        set myDate = new AvailableDate
        myDate.Load SelectedDate

        rowString = """rows"":["

        currentTime = myDate.StartTime()

        while currentTime <= myDate.EndTime
            rowString = rowString & "{""c"":[{""v"":""" & currentTime & """,""f"":null},"

            for i = 0 to SelectedEvents.Count - 1
                if SelectedEvents(i).IsPaid then
                    dim personCount

                    dim startDateTime, currentDateTime, endDateTime
                    startDateTime = SelectedEvents(i).SelectedDate & " " & GetMilitaryTime(SelectedEvents(i).StartTime)
                    currentDateTime = myDate.SelectedDate & " " & currentTime
                
                    if len(selectedEvents(i).CheckOutTime) > 0 then
                        endDateTime = SelectedEvents(i).SelectedDate & " " & GetMilitaryTime(SelectedEvents(i).CheckOutTime)
                    else
                        endDateTime = SelectedEvents(i).SelectedDate & " " & GetMilitaryTime(SelectedEvents(i).EstimatedEndTime)
                    end if

                    dim startDiff, endDiff
                    startDiff = DateDiff("n", startDateTime, currentDateTime)
                    endDiff = DateDiff("n", endDateTime, currentDateTime)

                    if startDiff >= 0 and endDiff <= 0 and len(SelectedEvents(i).NumberOfPatrons) > 0 then
                        personCount = SelectedEvents(i).NumberOfPatrons
                    else
                        personCount = 0
                    end if

                    rowString = rowString & "{""v"":" & personCount & ",""f"":null}"

                    if i < SelectedEvents.Count - 1 then
                        rowString = rowString & ","
                    end if
                end if
            next

            if currentTime < myDate.EndTime then
                rowString = rowString & "]},"
            else
                rowString = rowString & "]}],""p"":null}"
            end if

            currentTime = FormatDateTime(DateAdd("n", 30, GetTimeString(currentTime)), vbShortTime)
        wend

        BuildDailyStackedReportRows = rowString
    end function

    '***********************************************************
    ' build a list of active parties that can be checked out
    ' for the scheduled day
    '***********************************************************
    private sub BuildPartyList(SelectedDate)
        dim myEvents
        set myEvents = new ScheduledEventCollection
        myEvents.LoadByDateWithoutCheckout SelectedDate

        if myEvents.Count = 0 then
            response.Write "<p style=""color: red"">No parties scheduled for the selected date</p>" & vbCrLf
        else
            response.Write "<h3>Party List</h3>" & vbCrLf
            response.Write "<ul class=""partyList"">" & vbCrLf

            for i = 0 to myEvents.Count - 1
                if myEvents(i).IsPaid then
                    response.Write "<li>"

                    dim partyName
                    if len(myEvents(i).PartyName) = 0 then
                        partyName = myEvents(i).ContactName
                    else
                        partyName = Replace(myEvents(i).PartyName, """", "'")
                    end if

                    response.Write partyName

                    'if len(myEvents(i).GunSize) > 0 then
                    '    response.Write " (" & myEvents(i).GunSizeShort & ")"
                    'end if

                    Response.Write "</li>"
                end if
            next

            response.Write "</ul>" & vbCrLf
        end if
    end sub

    '**********************************************************
    ' Build an option list for the event types
    '**********************************************************
    sub BuildEventTypeOptions()
        'dim types
        'set types = new EventTypeCollection

        'types.LoadAllActive

        'for i = 0 to types.Count - 1
        '    dim myType
        '    set myType = types(i)

        '    response.Write "<option value='" & myType.EventTypeId & "'>" & myType.Description & "</option>"
        'next
    end sub

    '**********************************************************
    ' build an option list for available times
    '**********************************************************
    sub BuildTimeOptions(OpenTime, CloseTime, SelectedDate)
        dim currentTime, endTime
        currentTime = OpenTime
        endTime = FormatDateTime(DateAdd("h", -2, GetTimeString(CloseTime)), vbShortTime)

        while currentTime <= endTime
            dim myEvents
            set myEvents = new ScheduledEventCollection

            myEvents.LoadByDateAndTime SelectedDate, currentTime

            if myEvents.TotalPatrons <= CInt(Settings(SETTING_REGISTRATION_MAXPLAYERS))then
                response.Write "<option value='" & currentTime & "'>" & GetTimeString(currentTime) & "</option>"
            end if

            currentTime = FormatDateTime(DateAdd("n", CInt(Settings(SETTING_BOOKING_INTERVAL)), GetTimeString(currentTime)), vbShortTime)
        wend
    end sub

    '**********************************************************
    ' cancel the specified event and send an email if requested
    '**********************************************************
    sub CancelExistingEvent(EventId, StatusId, SendEmail)
        dim myEvent : set myEvent = new ScheduledEvent
        myEvent.Load EventId
        

        if cdbl(myEvent.DepositAmount) = 0.00 then
            myEvent.Delete EventId
        else
            myEvent.PaymentStatusId = StatusId
            myEvent.Save
        end if

        if SendEmail then 
            dim myEmail : set myEmail = new CancellationEmail
            myEmail.SendByObject myEvent
        end if
    end sub

    '**********************************************************
    ' set the checkout time for the selected event
    '**********************************************************
    sub CheckOutEvent(EventId, CheckOutTime)
        dim myEvent
        set myEvent = new ScheduledEvent
        myEvent.Load EventId
        myEvent.CheckOutTime = CheckOutTime
        myEvent.Save

        response.Write "Check out complete for: " & vbCrLf
        response.Write "Party: " & myEvent.PartyName & vbCrLf
        response.Write "Contact: " & myEvent.ContactName & vbCrLf
        response.Write "Time: " & CheckOutTime & vbCrLf
    end sub

    '**********************************************************
    ' see if there are scheduled events for the contact in the 
    ' future
    '**********************************************************
    sub CheckEventExistsByEmail(EmailAddress)
        dim myEvents : set myEvents = new ScheduledEventCollection
        myEvents.LoadByEmail EmailAddress
        
        Response.Clear
        Response.ContentType = MIMETYPE_JSON
        
        response.Write myEvents.ToJSON
    end sub

    function GetCartForEvent(CurrentEvent)
        dim details : set details = new OrderDetailCollection
        details.LoadByEventId CurrentEvent.EventId

        dim output
        dim adjustments : set adjustments = new ArrayList
        dim originalTotal : originalTotal = 0.00

        if details.Count = 0 then
            output = ""
        else
            output = output & "<table class='tblDetails'>"
                        
            for i = 0 to details.Count - 1
                if details(i).IsAdjustment then
                    adjustments.Add details(i).ItemNumber
                else
                    output = output & "<tr>"
                    output = output & "<td>" & details(i).ItemDescription & "</td>"
                    output = output & "<td>" & details(i).Quantity & "</td>"
                    output = output & "<td>" & FormatNumber(details(i).Price, 2) & "</td>"
                    output = output & "<td>" & FormatNumber(details(i).ExtendedPrice, 2) & "</td>"
                    output = output & "</tr>"

                    originalTotal = originalTotal + details(i).ExtendedPrice
                end if
            next

            output = output & "<tr>"
            output = output & "<td colspan='4'><hr/></td>"
            output = output & "</tr>"
            output = output & "<tr>"
            output = output & "<td colspan='3'>Total</td>"
            output = output & "<td>" & FormatNumber(originalTotal, 2) & "</td>"
            output = output & "</tr>"

            dim ps : set ps = new ScheduledEventPayment
            ps.LoadLastPaymentByEventId CurrentEvent.EventId

            if ps.PaymentStatus = PAYPAL_PAYMENTSTATUS_PAID then
                output = output & "<tr>"
                output = output & "<td colspan='4'>"
    
                if ps.PayPalTransactionId = PAYPAL_TRANSID_MANUAL then
                    output = output & "Paid by Admin - $" & FormatNumber(CurrentEvent.DepositAmount, 2) & " - " & CurrentEvent.ConfirmationDateTime
                else
                    output = output & "Paid by PayPal - $" & FormatNumber(ps.PaymentAmount, 2) & " - " & CurrentEvent.ConfirmationDateTime
                end if

                output = output & "</td>"
                output = output & "</tr>"
            else
                if cint(CurrentEvent.PaymentStatusId) < cint(PAYMENTSTATUS_CANCELLED_USER) then
                    output = output & "<tr>"
                    output = output & "<td colspan='4'><a href='javascript: markAsPaid(" & CurrentEvent.EventId & ");'>Mark as paid</a></td>"
                    output = output & "</tr>"
                end if
            end if


            if adjustments.Count > 0 then
                dim adjustedTotal : adjustedTotal = 0.00

                output = output & "<tr>"
                output = output & "<td colspan='4' class='center'><h3>Modified Order Details</h3></td>"
                output = output & "</tr>"

                for i = 0 to details.Count - 1
                    dim writableRecord : writableRecord = false

                    if details(i).IsAdjustment then
                        writableRecord = true
                    else
                        dim adjustmentFound : adjustmentFound = false

                        for j = 0 to adjustments.Count - 1
                            if details(i).ItemNumber = adjustments(j) then
                                adjustmentFound = true
                                j = adjustments.Count
                            end if
                        next

                        if adjustmentFound = false then writableRecord = true
                    end if

                    if writableRecord then
                        output = output & "<tr>"
                        output = output & "<td>" & details(i).ItemDescription & "</td>"
                        output = output & "<td>" & details(i).Quantity & "</td>"
                        output = output & "<td>" & FormatNumber(details(i).Price, 2) & "</td>"
                        output = output & "<td>" & FormatNumber(details(i).ExtendedPrice, 2) & "</td>"
                        output = output & "</tr>"

                        adjustedTotal = adjustedTotal + details(i).ExtendedPrice
                    end if
                next

                output = output & "<tr>"
                output = output & "<td colspan='4'><hr/></td>"
                output = output & "</tr>"
                output = output & "<tr>"
                output = output & "<td colspan='3'>Total</td>"
                output = output & "<td>" & FormatNumber(adjustedTotal, 2) & "</td>"
                output = output & "</tr>"

                if CurrentEvent.PaymentStatusId = PAYMENTSTATUS_PAID or CurrentEvent.PaymentStatusId = PAYMENTSTATUS_MARKEDPAID then
                else
                    dim differenceTotal : differenceTotal = originalTotal - adjustedTotal
                    output = output & "<tr>"
                    output = output & "<td colspan='3'>"
                    if differenceTotal = 0 then
                        output = output & "No action required"
                    elseif differenceTotal > 0 then
                        output = output & "We owe the customer"
                    else
                        output = output & "The customer owes us"
                    end if
                    output = output & "</td>"
                    output = output & "<td>" & FormatNumber(abs(differenceTotal), 2) & "</td>"
                    output = output & "</tr>"
                end if
            end if

            if CurrentEvent.IsCancelled then
                output = output & "<tr>"
                output = output & "<td colspan='4'>"
    
                select case CurrentEvent.PaymentStatusId
                    case PAYMENTSTATUS_CANCELLED_USER
                        output = output & "Canceled by user"
                    case PAYMENTSTATUS_CANCELLED_ADMIN
                        output = output & "Canceled by admin"
                    case PAYMENTSTATUS_CANCELLED_TIMEOUT
                        output = output & "Canceled by timeout"
                end select
                
                output = output & "</td>"
                output = output & "</tr>"
            end if

            output = output & "</table>"
        end if

        GetCartForEvent = output
    end function

    function GetPaymentText(PaymentTypeId, Amount)
        dim text
        text = "$" 
        text = text & FormatNumber(cdbl(Amount), 2)

        select case cstr(PaymentTypeId)
            case PAYMENTTYPE_BYEVENT
                'text = text & "event"
            case PAYMENTTYPE_BYPLAYER
                text = text & " / person"
        end select
        
        GetPaymentText = text
    end function

    sub MarkAsPaid(EventId)
        dim myEvent : set myEvent = new ScheduledEvent
        myEvent.Load EventId

        myEvent.PaymentStatusId = PAYMENTSTATUS_MARKEDPAID
        myEvent.ConfirmationDateTime = Now()
        myEvent.Save

        dim payment : set payment = new ScheduledEventPayment
        payment.EventId = EventId
        payment.StatusDateTime = Now()
        payment.PaymentAmount = myEvent.DepositAmount
        payment.PaymentStatus = PAYPAL_PAYMENTSTATUS_PAID
        payment.PayPalTransactionId = PAYPAL_TRANSID_MANUAL
       
        payment.Save

        dim details : set details = new OrderDetailCollection
        details.LoadByEventId EventId

        for i = 0 to details.Count - 1
            details(i).PaymentTransaction = PAYPAL_TRANSID_MANUAL
            details(i).PaymentDateTime = Now()
            details(i).PaymentAmount = details(i).ExtendedPrice
        next

        details.SaveAll

        dim myEmail : set myEmail = new ConfirmationEmail
        myEmail.SendByObject myEvent
    end sub

    '**********************************************************
    ' build the form for registration
    '**********************************************************
    sub ShowRegistrationForm(EncodedEventId)
        dim myEvent, myDate
        set myEvent = new ScheduledEvent
        set myDate = new AvailableDate

        myEvent.Load DecodeId(EncodedEventId)
        myDate.Load(myEvent.SelectedDate)

        dim eventDiff : daysRemaining = DateDiff("d", Now(), myDate.ToDate())
    
        with Response
            if myEvent.Active = false then
                .Write "<h1 class=""center"">Event Cancelled</h1>"
                .Write "<p class=""center"">This event has been cancelled. You can <a href=""default.asp"" rel=""external"">book another event.</a></p>"
                .Write "<p class=""center""><a href=""" & SiteInfo.HomeUrl & """ rel=""external"">Return home</a></p>"
            elseif Now() > myDate.ToDate() then
                .Write "<h1 class=""center"">Event In The Past</h1>"
                .Write "<p class=""center"">This scheduled time for this event has passed. You can <a href=""default.asp"" rel=""external"">book another event.</a></p>"
                .Write "<p class=""center""><a href=""" & SiteInfo.HomeUrl & """ rel=""external"">Return home</a></p>"
            elseif (cbool(Settings(SETTING_REGISTRATION_ALLOWCANCELLATION)) and daysRemaining < cint(Settings(SETTING_REGISTRATION_MAXDAYS_CANCEL))) and _
                   (cbool(Settings(SETTING_REGISTRATION_ALLOWEDIT)) and daysRemaining < cint(Settings(SETTING_REGISTRATION_MAXDAYS_EDIT))) then
                .Write "<h1 class=""center"">Event Cannot Be Modified</h1>"
                .Write "<p class=""center"">The allowed time to cancel or change this event has passed. You can <a href=""default.asp"" rel=""external"">book another event.</a></p>"
                .Write "<p class=""center""><a href=""" & SiteInfo.HomeUrl & """ rel=""external"">Return home</a></p>"
            elseif (cbool(Settings(SETTING_REGISTRATION_ALLOWCANCELLATION)) and daysRemaining < cint(Settings(SETTING_REGISTRATION_MAXDAYS_CANCEL))) and _
                    cbool(Settings(SETTING_REGISTRATION_ALLOWEDIT)) = false then
                .Write "<h1 class=""center"">Event Cannot Be Canceled</h1>"
                .Write "<p class=""center"">The allowed time to cancel this event has passed. You can <a href=""default.asp"" rel=""external"">book another event.</a></p>"
                .Write "<p class=""center""><a href=""" & SiteInfo.HomeUrl & """ rel=""external"">Return home</a></p>"
            elseif (cbool(Settings(SETTING_REGISTRATION_ALLOWEDIT)) and daysRemaining < cint(Settings(SETTING_REGISTRATION_MAXDAYS_EDIT))) and _
                    cbool(Settings(SETTING_REGISTRATION_ALLOWCANCELLATION)) = false then
                .Write "<h1 class=""center"">Event Cannot Be Modified</h1>"
                .Write "<p class=""center"">The allowed time to modify this event has passed. You can <a href=""default.asp"" rel=""external"">book another event.</a></p>"
                .Write "<p class=""center""><a href=""" & SiteInfo.HomeUrl & """ rel=""external"">Return home</a></p>"
            elseif cbool(Settings(SETTING_REGISTRATION_ALLOWCANCELLATION)) = false and cbool(Settings(SETTING_REGISTRATION_ALLOWEDIT)) = false then
                .Write "<h1 class=""center"">Event Cannot Be Modified</h1>"
                .Write "<p class=""center"">This site has decided not to allow cancellations or modifications to events. You can <a href=""default.asp"" rel=""external"">book another event.</a></p>"
                .Write "<p class=""center""><a href=""" & SiteInfo.HomeUrl & """ rel=""external"">Return home</a></p>"
            else
                .Write "<img src=""" & LOGO_URL & """ style=""display: block; margin: auto"" />"
                .Write "<h1 class=""center"">Update Event Registration</h1>"
                .Write "<h3 class=""center"">Please Completely Fill Out The Registration Form</h3>"
                .Write "<table cellpadding=""2"" cellspacing=""3"" border=""0"" style=""margin: 0 auto;"">"
                .Write "<tr>"
                .Write "<td style=""width:145px; text-align: right"">Date:</td>"
                .Write "<td>"
                .Write "<input type=""text"" name=""SelectedDate"" id=""SelectedDate"" value=""" & FormatDateTime(myEvent.SelectedDate, vbShortDate) & """ readonly=""readonly"" data-inline=""true"" />"
                .Write "</td>"
                .Write "<td>"
                .Write "<img id=""btnCalendar"" src=""../content/images/calendar_view_month.png"" alt=""calendar"" class=""tooltip"" data-powertip=""If you are changing the date, the time you had picked previously may not be available.  After you pick your date, the available times will be present in the drop down below"" />"
                .Write "</td>"
                .Write "<td/>"
                .Write "</tr>"
                .Write "<tr>"
                .Write "<td style=""width:145px; text-align: right"">Your Name:</td>"
                .Write "<td>"
                .Write "<input type=""text"" name=""Name"" maxlength=""100"" onblur=""validateContactName()"" value=""" & myEvent.ContactName & """ /><br />"
                .Write "<span name=""ValName"" style=""display: none; color: Red"">Name is required</span>"
                .Write "</td>"
                .Write "<td style=""color: Red""> &nbsp; *</td>"
                .Write "<td/>"
                .Write "</tr>"
                .Write "<tr>"
                .Write "<td style=""width:145px; text-align: right"">Your Email Address:</td>"
                .Write "<td>"
                .Write "<input type=""email"" name=""Email"" maxlength=""128"" onblur=""validateEmail()"" value=""" & myEvent.ContactEmailAddress & """ />"
                .Write "<span name=""ValEmail"" style=""display: none; color: Red"">Invalid email address</span>"
                .Write "<span name=""ValEmailRequired"" style=""display: none; color: Red"">Email address required</span>"
                .Write "</td>"
                .Write "<td style=""color: Red""> &nbsp; *</td>"
                .Write "<td/>"
                .Write "</tr>"
                .Write "<tr style=""display: none"">"
                .Write "<td style=""width:145px; text-align: right"">Please Confirm Email:</td>"
                .Write "<td>"
                .Write "<input type=""email"" name=""ConfirmEmail"" maxlength=""128"" onblur=""validateEmail()"" value=""" & myEvent.ContactEmailAddress & """ />"
                .Write "<span name=""ValEmailMatch"" style=""display: none; color: Red"">Email address does not match</span>"
                .Write "</td>"
                .Write "<td colspan=""2""/>"
                .Write "</tr>"
                .Write "<tr>"
                .Write "<td style=""width:145px; text-align: right"">Approximate Number Of Guests:</td>"
                .Write "<td>"
                .Write "<input type=""text"" name=""NumberOfGuests"" id=""NumberOfGuests"" maxlength=""3"" value=""" & myEvent.NumberOfPatrons & """ onblur=""this.value = this.value.trim(); NumberOfGuests_onblur()"" />"
                .Write "<span name=""ValNumberOfGuests"" style=""display: none; color: Red"">Invalid number</span>"
                .Write "<span name=""ValNumberOfGuestsRequired"" style=""display: none; color: Red"">Number of guests required</span>"
                .Write "</td>"
                .Write "<td style=""color: Red""> &nbsp; *</td>"
                .Write "<td/>"
                .Write "</tr>"
                .Write "<tr>"
                .Write "<td style=""width:145px; text-align: right"">Hours of Operation:</td>"
                .Write "<td colspan=""3"">"
                .Write "<span name=""openHours""></span>"
                .Write "</td>"
                .Write "</tr>"
                .Write "<tr>"
                .Write "<td style=""width:145px; text-align: right"">Start Time:</td>"
                .Write "<td>"
                .Write "<table style=""width: 100%"">"
                .Write "<tr>"
                .Write "<td>"
                .Write "<input type=""hidden"" name=""Time"" id=""Time"" value=""" & myEvent.StartTimeShortFormat & """ />"
                .Write "<span id=""lblStartTime"">" & GetTimeString(myEvent.StartTime) & "</span>"
                .Write "</td>"
                .Write "<td class=""right"" style=""float: right"">"
                .Write "<button data-rel=""popup"" onclick=""return showAvailableTimes()"" id=""btnTime"" data-inline=""true"" data-mini=""true"">Choose Time</button>"
                .Write "</td>"
                .Write "</tr>"
                .Write "</table>"
                .Write "<span name=""ValTime"" style=""display: none; color: Red"">Start time is required</span>"
                .Write "</td>"
                .Write "<td style=""color: Red""> &nbsp; *</td>"
                .Write "<td/>"
                .Write "</tr>"
                .Write "<tr>"
                .Write "<td style=""width:145px; text-align: right"">Average Age Of Guests:</td>"
                .Write "<td>"
                .Write "<input type=""text"" name=""AgeOfGuests"" maxlength=""100"" value=""" & myEvent.AgeOfPatrons & """ />"
                .Write "</td>"
                .Write "<td colspan=""2""/>"
                .Write "</tr>"
                .Write "<tr>"
                .Write "<td style=""width:145px; text-align: right"">Name For Your Party:</td>"
                .Write "<td>"
                .Write "<input type=""text"" name=""PartyName"" maxlength=""100"" value=""" & myEvent.PartyName & """ class=""tooltip"" data-powertip=""Note: Please use a meaningful name for the party for your guests to find easily while filling out waivers.""/>"
                .Write "</td>"
                .Write "<td colspan=""2"" />"
                .Write "</tr>"
                .Write "<tr>"
                .Write "<td style=""width:145px; text-align: right"">Contact Phone Number:</td>"
                .Write "<td>"
                .Write "<input type=""tel"" name=""Phone"" maxlength=""20"" onblur=""validatePhoneNumber()"" value=""" & myEvent.ContactPhone & """ />"
                .Write "<span name=""ValPhoneRequired"" style=""display: none; color: Red"">Phone number required</span>"
                .Write "<span name=""ValPhoneFormat"" style=""display: none; color: Red"">Invalid phone number</span>"
                .Write "</td>"
                .Write "<td style=""color: Red""> &nbsp; *</td>"
                .Write "<td/>"
                .Write "</tr>"           
                
                BuildCustomFields myEvent.EventId

                .Write "<tr>"
                .Write "<td colspan=""4"">" & Settings(SETTING_BLURB_REGISTRATION)
                .Write "</td>"
                .Write "</tr>"                
                .Write "<tr>"
                .Write "<td style=""width:145px; text-align: right"">Comments:</td>"
                .Write "<td colspan=""3"">"
                .Write "<span name=""userCharCounter"">(1000 characters available)</span>"
                .Write "</td>"
                .Write "</tr>"
                .Write "<tr>"
                .Write "<td colspan=""4"">"
                .Write "<textarea name=""UserComments"" style=""resize: none; width: 98%"" maxlength=""1000"" onkeyup=""handleUserMaxLength()"">" & myEvent.UserComments & "</textarea>"
                .Write "</td>"
                .Write "</tr>"
                .Write "<tr>"
                .Write "<td colspan=""4"">"
                .Write "<input type=""hidden"" id=""txtBaseFee"" value=""" & FormatNumber(CDbL(Settings(SETTING_REGISTRATION_PAYMENTAMOUNT)),2) & """ />"
                .Write "<input type=""hidden"" id=""txtBaseFeeType"" value=""" & Settings(SETTING_REGISTRATION_PAYMENTTYPE) & """ />"
                .Write "<input type=""hidden"" id=""txtCartTotal"" />"
                .Write "<table id=""tblPaymentDetails"" class=""blockCenter"">"
                .Write "<tr>"
                .Write "<th>Description</th>"
                .Write "<th>Qty</th>"
                .Write "<th>Price</th>"
                .Write "<th>Ext. Price</th>"
                .Write "</tr>"
                .Write "<tr>"
                .Write "<td colspan=""4"">"
                .Write "<hr/>"
                .Write "</td>"
                .Write "</tr>"
                .Write "<tr>"
                .Write "<td colspan=""3"" />"
                .Write "<td class=""inlineRight"">0.00</td>"
                .Write "</tr>"
                .Write "</table>"
                .Write "</td>"
                .Write "</tr>"
                .Write "<tr>"
                .Write "<td colspan=""4"" style=""text-align: center"">"
                .Write "<div data-role=""controlgroup""  data-type=""horizontal"">"
                if cbool(Settings(SETTING_REGISTRATION_ALLOWEDIT)) and daysRemaining => cint(Settings(SETTING_REGISTRATION_MAXDAYS_EDIT)) then
                    .Write "<button id=""btnSubmit"" type=""submit"" onclick=""return validateForm();"" ondblclick=""return false;"" data-theme=""b"">Register</button>"
                end if
                .Write "<button type=""button"" onclick=""return redirectToHome();"">Cancel</button>"
                if cbool(Settings(SETTING_REGISTRATION_ALLOWCANCELLATION)) and daysRemaining => cint(Settings(SETTING_REGISTRATION_MAXDAYS_CANCEL)) then
                    .Write "<button onclick=""return cancelRegistration()"">Cancel This Reservation</button>"
                end if
                .Write "</div>"
                .Write "</td>"
                .Write "</tr>"
                .Write "</table>"
                .Write "<div data-role=""popup"" id=""divTimes"">"
                .Write "<iframe style=""border: none; padding: 25px""></iframe>"
                .Write "</div>"
                .Write "<script type=""text/javascript"">"
                '.Write "initializeForm('" & myEvent.EventTypeId & "', '" & myEvent.NumberOfPatrons & "', '" & GetMilitaryTime(myEvent.StartTime) & "');"
                '.Write "initializeForm('" & myEvent.NumberOfPatrons & "', '" & GetMilitaryTime(myEvent.StartTime) & "');"
                .Write "initializeForm();"
                .Write "</script>"
            end if
        end with
    end sub

    '**********************************************************
    ' build the registration validation form
    '**********************************************************
    sub ShowRegistrationValidationForm(EmailAddress, FirstAttempt)
        with Response
            .Write "<img src=""" & LOGO_URL & """ style=""display: block; margin: auto"" />"
            .Write "<h3 class=""center"">Registration Verification</h3>"
            .Write "<p class=""center"">Please enter the email address used to register your party. <br />(Hint: use the email address that received the confirmation email)</p>"
            .Write "<table cellpadding=""3"" cellspacing=""2"" border=""0"" style=""margin: 0 auto"">"
            .Write "<tr style=""vertical-align: top"">"
            .Write "<td>Email Address</td>"
            .Write "<td>"
            .Write "<input type=""text"" name=""Email"" value=""" & EmailAddress & """ /><br />"
            .Write "<span name=""ValEmailRequired"" style=""color: Red; display: none"">Required</span>"
            .Write "<span name=""ValEmail"" style=""color: Red; display: none"">Invalid address</span>"
            if FirstAttempt = false then
                .Write "<span style=""color: Red;"">Email does not match</span>"
            end if
            .Write "</td>"
            .Write "<td style=""color: Red""> &nbsp; *</td>"
            .Write "</tr>"
            .Write "<tr>"
            .Write "<td colspan=""3"" class=""center"">"
            .Write "<button name=""btnSubmit"" onclick=""submitVerification(); return false;"">Submit</button>"
            .Write "</td>"
            .Write "</tr>"
            .Write "</table>"
        end with
    end sub

    '**********************************************************
    ' verify that the email address matches what is saved on the event
    '**********************************************************
    function VerifyEmail(EncodedEventId, EmailAddress)
        dim myEvent
        set myEvent = new ScheduledEvent
        myEvent.Load DecodeId(EncodedEventId)

        if myEvent.Loaded then
            VerifyEmail = CBool(StrComp(myEvent.ContactEmailAddress, EmailAddress, 1) = 0)
        else
            VerifyEmail = false
        end if
    end function
        
    '**********************************************************
    ' build the list of events for the specified date
    '**********************************************************
    sub WriteEventTableForDay(ScheduledDate)
        dim usesDeposits : usesDeposits = UsesPayments()
        dim COLUMNS_IN_ROW : COLUMNS_IN_ROW = iif(usesDeposits, 8, 7)

        dim events, statusDict
        set events = new ScheduledEventCollection
        events.LoadByDate ScheduledDate
        
        set statusDict = Server.CreateObject("Scripting.Dictionary")
        
        '*********************************************************
        ' this format does not match the format from the 
        ' next method. cannot refactor to reuse code
        '*********************************************************
        response.Write "<table id=""scheduleTable"" cellpadding=""2"" cellspacing=""0"" border=""1"">"
        response.Write "<thead>"
        response.Write "<tr class=""heading"">"
        response.Write "<th nowrap=""nowrap"">Party Name</th>"
        response.Write "<th>Contact</th>"
        if usesDeposits then response.Write "<th>Status</th>"
        'response.Write "<th>Event Type</th>"
        response.Write "<th>Number</th>"
        response.Write "<th>Waivers</th>"
        response.Write "<th>Age</th>"
        response.Write "<th>Start</th>"
        response.Write "<th>Created</th>"
        response.Write "</tr>"
        response.Write "</thead>"
        response.Write "<tbody>"

        for i = 0 to events.Count - 1
            dim myEvent
            set myEvent = events(i)

            response.Write "<tr>"
            'response.Write "<td>" & myEvent.EventId & "</td>"
            response.Write "<td><img src=""../../content/images/pencil.png"" alt="""" title=""Edit"" border=""0"" class=""mouseover"" onclick=""editRegistration(" & myEvent.EventId & ")"" />" & myEvent.PartyName & "</td>"
            response.Write "<td>" & myEvent.ContactName
            response.Write "<img src=""../../content/images/note.png"" alt="""" data-powertip=""Email: <a href='mailto:" & myEvent.ContactEmailAddress & "'>" & myEvent.ContactEmailAddress & "</a>" & _
                            "<br/>Phone: " & myEvent.ContactPhone  & """ border=""0"" class=""tooltip"" style=""float: right""/>"
            response.Write "</td>"
            
            if usesDeposits then
                dim bPaid, sImgUrl, sToolTip, textColor
                select case lcase(myEvent.PaymentStatusId)
                    case PAYMENTSTATUS_PAID
                        textColor = "green"
                        bPaid = true
                        sImgUrl = "../../content/images/tick.png"
                        sToolTip = "paid"
                    case PAYMENTSTATUS_MARKEDPAID
                        textColor = "green"
                        bPaid = true
                        sImgUrl = "../../content/images/tick.png"
                        sToolTip = "paid"
                    case PAYMENTSTATUS_CHANGEDPAYABLE
                        textColor = "orange"
                        sImgUrl = "../../content/images/error.png"
                        bPaid = true
                        sToolTip = "Changed - Payable"
                    case PAYMENTSTATUS_CHANGEDRECEIVABLE
                        textColor = "orange"
                        sImgUrl = "../../content/images/error.png"
                        bPaid = true
                        sToolTip = "Changed - Receivable"
                    case PAYMENTSTATUS_CANCELLED_USER, PAYMENTSTATUS_CANCELLED_ADMIN, PAYMENTSTATUS_CANCELLED_TIMEOUT
                        textColor = "black"

                        dim ps : set ps = new ScheduledEventPayment
                        ps.LoadLastPaymentByEventId myEvent.EventId

                        if ps.PaymentStatus = PAYPAL_PAYMENTSTATUS_PAID then
                            sImgUrl = "../../content/images/nomoney.png"
                        else
                            sImgUrl = "../../content/images/no.png"
                        end if

                        bPaid = false
                        sToolTip = "Cancelled"
                    case else
                        textColor = "red"
                        sImgUrl = "../../content/images/exclamation.png"
                        sToolTip = "unpaid"
                        bPaid = false
                end select
                response.Write "<td style=""color: " & textColor & """>" '& status.PaymentStatus & "</td>"
                
                dim cartText : cartText = GetCartForEvent(myEvent)
                if len(cartText) > 0 then
                    response.Write "<img src=""../../content/images/cart_empty.png"" data-powertip=""" & cartText & """ class=""tooltip"" />"
                end if
                Response.Write "<img src=""" & sImgUrl & """ alt=""" & sToolTip & """ title=""" & sToolTip & """ />"
                response.Write FormatNumber(myEvent.DepositAmount, 2)
                response.Write "</td>"

                if bPaid then
                    if len(statusDict(KEY_PAYMENTSTATUS_PAID)) = 0 then
                        statusDict(KEY_PAYMENTSTATUS_PAID) = "1;" & myEvent.NumberOfPatrons
                    else
                        dim paidNumbers : paidNumbers = split(statusDict(KEY_PAYMENTSTATUS_PAID), ";")
                        statusDict(KEY_PAYMENTSTATUS_PAID) = cstr(cint(paidNumbers(0)) + 1) & ";" & cstr(cint(paidNumbers(1)) + myEvent.NumberOfPatrons)
                    end if
                else
                    if len(statusDict(KEY_PAYMENTSTATUS_UNPAID)) = 0 then
                        statusDict(KEY_PAYMENTSTATUS_UNPAID) = "1;" & myEvent.NumberOfPatrons
                    else
                        dim unpaidNumbers : unpaidNumbers = split(statusDict(KEY_PAYMENTSTATUS_UNPAID), ";")
                        statusDict(KEY_PAYMENTSTATUS_UNPAID) = cstr(cint(unpaidNumbers(0)) + 1) & ";" & cstr(cint(unpaidNumbers(1)) + myEvent.NumberOfPatrons)
                    end if
                end if
            end if

            response.Write "<td>" & myEvent.NumberOfPatrons & "</td>"
            'if myEvent.WaiverCount = 0 then
            '    response.Write "<td>" & myEvent.WaiverCount & "</td>"
            'else
            '    response.Write "<td><a href=""../waiver/waiverlist.asp?ev=" & myEvent.EventId & "&dt=" & myEvent.SelectedDate & """ target=""_blank"">" & myEvent.WaiverCount & "</a></td>"
            'end if
            response.Write "<td><a href=""../waiver/default.asp?id=" & myEvent.EventId & "&dt=" & myEvent.SelectedDate & """>" & myEvent.WaiverCount & "/" & myEvent.NumberOfPatrons & "</a></td>"
            response.Write "<td>" & myEvent.AgeOfPatrons & "</td>"
            response.Write "<td>" & myEvent.StartTimeLongFormat & "</td>"
            response.Write "<td>" & myEvent.CreateDateTime & "</td>"
            response.Write "</tr>"
    
            '***************************************************************
            ' custom fields
            '***************************************************************
            myEvent.LoadValues
            dim customValues : set customValues = myEvent.CustomFieldValues
            if customValues.Count > 0 then
                dim shortValues, longValues
                set shortValues = new RegistrationCustomValueCollection
                set longValues = new RegistrationCustomValueCollection

                for j = 0 to myEvent.CustomFieldValues.Count - 1
                    dim myValue : set myValue = customValues(j)
                    dim myField : set myField = new RegistrationCustomField
                    
                    myField.LoadById myValue.RegistrationCustomFieldId
                    
                    if myField.CustomFieldDataType.CustomFieldDataTypeId = FIELDTYPE_LONGTEXT then
                        longValues.Add myValue
                    else
                        shortValues.Add myValue
                    end if
                next

                dim colsUsed
                for shortValueIndex = 0 to shortValues.Count - 1
                    if shortValueIndex = 0 then
                        response.Write "<tr class=""expand-child"">"
                        colsUsed = 0
                    elseif shortValueIndex mod 3 = 0 then
                        if COLUMNS_IN_ROW - colsUsed > 0 then response.Write "<td colspan=""" & COLUMNS_IN_ROW - colsUsed & """/>"
                        response.Write "</tr>"
                        response.Write "<tr class=""expand-child"">"
                        colsUsed = 0
                    end if

                    dim ShortField : set ShortField = new RegistrationCustomField
                    ShortField.LoadById shortValues(shortValueIndex).RegistrationCustomFieldId

                    response.Write "<td class=""heading"">" & ShortField.Name &"</td>"
        
                    if ShortField.CustomFieldDataType.CustomFieldDataTypeId = FIELDTYPE_OPTIONS then
                        dim optionValue : set optionValue = new RegistrationCustomOption
                        optionValue.RegistrationCustomFieldId = ShortField.RegistrationCustomFieldId
                        optionValue.Sequence = shortValues(shortValueIndex).Value
                        optionValue.Load

                        Response.Write "<td colspan=""2"">" & optionValue.Text & "</td>"
                    else
                        Response.Write "<td colspan=""2"">" & shortValues(shortValueIndex).Value & "</td>"
                    end if

                    colsUsed = colsUsed + 3
                next

                if shortValues.Count > 0 then 
                    if COLUMNS_IN_ROW - colsUsed > 0 then response.Write "<td colspan=""" & COLUMNS_IN_ROW - colsUsed & """/>"
                    response.Write "</tr>"
                end if

                for longValueIndex = 0 to longValues.Count - 1
                    dim longField : set longField = new RegistrationCustomField
                    longField.LoadById longValues(longValueIndex).RegistrationCustomFieldId

                    response.Write "<tr class=""expand-child"">"
                    response.Write "<td class=""heading"" nowrap=""nowrap"">" & longField.Name & "</td>"
                    Response.Write "<td colspan=""" & COLUMNS_IN_ROW - 1 & """>" & longValues(longValueIndex).Value & "</td>"
                    response.Write "</tr>"
                next
            end if        

            response.Write "<tr class=""expand-child"">"
            response.Write "<td class=""heading"" nowrap=""nowrap"">User Notes</td>"
            Response.Write "<td colspan=""" & COLUMNS_IN_ROW - 1 & """>" & myEvent.UserComments & "</td>"
            response.Write "</tr>"

            response.Write "<tr class=""expand-child"">"
            response.Write "<td class=""heading"" nowrap=""nowrap"">Admin Notes</td>"
            Response.Write "<td colspan=""" & COLUMNS_IN_ROW - 1 & """>" & myEvent.AdminComments & "</td>"
            response.Write "</tr>"

            response.Write "<tr class=""expand-child"">"
            Response.Write "<td colspan=""" & COLUMNS_IN_ROW & """ style=""background-color: grey"">&nbsp;</td>"
            response.Write "</tr>"
        next

        response.Write "</tbody>"
        
        '****************************************************
        ' footer
        '****************************************************
        Response.Write "<tfoot>"

        dim numberTitleColWritten : numberTitleColWritten = false

        for each statusKey in statusDict
            dim totalNumbers
            totalNumbers = split(statusDict(statusKey), ";")

            response.Write "<tr>"
            Response.Write "<td class=""heading"">" & statusKey & " Totals</td>"
            if numberTitleColWritten = false then
                Response.Write "<td colspan=""2"">" & totalNumbers(0) & "</td>"
                Response.Write "<td class=""heading"">Number of guests</td>"
                numberTitleColWritten = true
            else
                Response.Write "<td colspan=""3"">" & totalNumbers(0) & "</td>"
            end if
            Response.Write "<td colspan=""5"">" & totalNumbers(1) & "</td>"
            response.Write "</tr>"
        next

        response.Write "<tr>"
        Response.Write "<td class=""heading"">Totals</td>"
        if numberTitleColWritten = false then
            Response.Write "<td colspan=""2"">" & events.Count & "</td>"
            Response.Write "<td class=""heading"">Number of guests</td>"
            numberTitleColWritten = true
        else
            Response.Write "<td colspan=""3"">" & events.Count & "</td>"
        end if
        if usesDeposits then
            Response.Write "<td colspan=""5"">" & events.TotalPatrons & "</td>"
        else
            Response.Write "<td colspan=""4"">" & events.TotalPatrons & "</td>"
        end if
        response.Write "</tr>"
        Response.Write "</tfoot>"
        
        response.Write "</table>"
    end sub

    '**********************************************************
    ' build the event table for all events from today and forward
    '**********************************************************
    sub WriteEventTableRemaining()
        dim usesDeposits : usesDeposits = UsesPayments()
        dim COLUMNS_IN_ROW : COLUMNS_IN_ROW = iif(usesDeposits, 9, 8)
        
        'dim events, types, myDict
        dim events
        set events = new ScheduledEventCollection
        'set types = new EventTypeCollection
        events.LoadRemaining
        'types.LoadAll

        'set myDict = types.ToDictionary()
        set statusDict = Server.CreateObject("Scripting.Dictionary")

        '*********************************************************
        ' this format does not match the format from the 
        ' previous method. cannot refactor to reuse code
        '*********************************************************
        response.Write "<table id=""scheduleTable"" cellpadding=""2"" cellspacing=""0"" border=""1"">"
        response.Write "<thead>"
        response.Write "<tr class=""heading"">"
        response.Write "<th nowrap=""nowrap"">Party Name</th>"
        response.Write "<th>Contact</th>"
        if usesDeposits then response.Write "<th>Status</th>"
        'response.Write "<th>Event Type</th>"
        response.Write "<th>Number</th>"
        response.Write "<th>Waivers</th>"
        response.Write "<th>Age</th>"
        response.Write "<th>Scheduled</th>"
        response.Write "<th>Start</th>"
        response.Write "<th>Created</th>"
        response.Write "</tr>"
        response.Write "</thead>"
        response.Write "<tbody>"

        for i = 0 to events.Count - 1
            dim myEvent
            set myEvent = events(i)

            response.Write "<tr>"
            'response.Write "<td>" & myEvent.EventId & "</td>"
            response.Write "<td><img src=""../../content/images/pencil.png"" alt="""" title=""Edit"" border=""0"" class=""mouseover"" onclick=""editRegistration(" & myEvent.EventId & ")"" />" & myEvent.PartyName & "</td>"
            response.Write "<td>" & myEvent.ContactName
            response.Write "<img src=""../../content/images/note.png"" alt="""" data-powertip=""Email: <a href='mailto:" & myEvent.ContactEmailAddress & "'>" & myEvent.ContactEmailAddress & "</a>" & _
                            "<br/>Phone: " & myEvent.ContactPhone  & """ border=""0"" class=""tooltip"" style=""float: right""/>"
            response.Write "</td>"
            'response.Write "<td>" & myEvent.ContactEmailAddress & "</td>"
            'response.Write "<td>" & myEvent.ContactPhone & "</td>"
            
            if usesDeposits then
                dim status : set status = new ScheduledEventPayment
                status.EventId = myEvent.EventId
                status.LoadLastPayment

                dim bPaid, sImgUrl, sToolTip, textColor
                select case lcase(myEvent.PaymentStatusId)
                    case PAYMENTSTATUS_PAID
                        textColor = "green"
                        bPaid = true
                        sImgUrl = "../../content/images/tick.png"
                        sToolTip = "paid"
                    case PAYMENTSTATUS_MARKEDPAID
                        textColor = "green"
                        bPaid = true
                        sImgUrl = "../../content/images/tick.png"
                        sToolTip = "paid"
                    case PAYMENTSTATUS_CHANGEDPAYABLE
                        textColor = "orange"
                        sImgUrl = "../../content/images/error.png"
                        bPaid = true
                        sToolTip = "Changed - Payable"
                    case PAYMENTSTATUS_CHANGEDRECEIVABLE
                        textColor = "orange"
                        sImgUrl = "../../content/images/error.png"
                        bPaid = true
                        sToolTip = "Changed - Receivable"
                    case PAYMENTSTATUS_CANCELLED_USER, PAYMENTSTATUS_CANCELLED_ADMIN, PAYMENTSTATUS_CANCELLED_TIMEOUT
                        textColor = "black"

                        dim ps : set ps = new ScheduledEventPayment
                        ps.LoadLastPaymentByEventId myEvent.EventId

                        if ps.PaymentStatus = PAYPAL_PAYMENTSTATUS_PAID then
                            sImgUrl = "../../content/images/nomoney.png"
                        else
                            sImgUrl = "../../content/images/no.png"
                        end if

                        bPaid = false
                        sToolTip = "Cancelled"
                    case else
                        textColor = "red"
                        sImgUrl = "../../content/images/exclamation.png"
                        sToolTip = "unpaid"
                        bPaid = false
                end select
                response.Write "<td style=""color: " & textColor & """>" '& status.PaymentStatus & "</td>"
                dim cartText : cartText = GetCartForEvent(myEvent)
                if len(cartText) > 0 then
                    response.Write "<img src=""../../content/images/cart_empty.png"" data-powertip=""" & cartText & """ class=""tooltip"" />"
                end if
                Response.Write "<img src=""" & sImgUrl & """ alt=""" & sToolTip & """ title=""" & sToolTip & """ />"
                response.Write FormatNumber(myEvent.DepositAmount, 2)
                response.Write "</td>"

                if bPaid then
                    if len(statusDict(KEY_PAYMENTSTATUS_PAID)) = 0 then
                        statusDict(KEY_PAYMENTSTATUS_PAID) = "1;" & myEvent.NumberOfPatrons
                    else
                        dim paidNumbers : paidNumbers = split(statusDict(KEY_PAYMENTSTATUS_PAID), ";")
                        statusDict(KEY_PAYMENTSTATUS_PAID) = cstr(cint(paidNumbers(0)) + 1) & ";" & cstr(cint(paidNumbers(1)) + myEvent.NumberOfPatrons)
                    end if
                else
                    if len(statusDict(KEY_PAYMENTSTATUS_UNPAID)) = 0 then
                        statusDict(KEY_PAYMENTSTATUS_UNPAID) = "1;" & myEvent.NumberOfPatrons
                    else
                        dim unpaidNumbers : unpaidNumbers = split(statusDict(KEY_PAYMENTSTATUS_UNPAID), ";")
                        statusDict(KEY_PAYMENTSTATUS_UNPAID) = cstr(cint(unpaidNumbers(0)) + 1) & ";" & cstr(cint(unpaidNumbers(1)) + myEvent.NumberOfPatrons)
                    end if
                end if
            end if

            'response.Write "<td>" 
            'if len(myEvent.EventTypeId) > 0 then response.Write myDict.Item(myEvent.EventTypeId) 
            'response.Write "</td>"
            response.Write "<td>" & myEvent.NumberOfPatrons & "</td>"
            'if myEvent.WaiverCount = 0 then
            '    response.Write "<td>" & myEvent.WaiverCount & "</td>"
            'else
            '    response.Write "<td><a href=""../waiver/waiverlist.asp?ev=" & myEvent.EventId & "&dt=" & myEvent.SelectedDate & """ target=""_blank"">" & myEvent.WaiverCount & "</a></td>"
            'end if
            response.Write "<td><a href=""../waiver/default.asp?id=" & myEvent.EventId & "&dt=" & myEvent.SelectedDate & """>" & myEvent.WaiverCount & "/" & myEvent.NumberOfPatrons & "</a></td>"
            response.Write "<td>" & myEvent.AgeOfPatrons & "</td>"
            response.Write "<td>" & myEvent.SelectedDate & "</td>"
            response.Write "<td>" & myEvent.StartTimeLongFormat & "</td>"
            response.Write "<td>" & myEvent.CreateDateTime & "</td>"
            response.Write "</tr>"
    
            '***************************************************************
            ' custom fields
            '***************************************************************
            myEvent.LoadValues
            dim customValues : set customValues = myEvent.CustomFieldValues
            if customValues.Count > 0 then
                dim shortValues, longValues
                set shortValues = new RegistrationCustomValueCollection
                set longValues = new RegistrationCustomValueCollection

                for j = 0 to myEvent.CustomFieldValues.Count - 1
                    dim myValue : set myValue = customValues(j)
                    dim myField : set myField = new RegistrationCustomField
                    
                    myField.LoadById myValue.RegistrationCustomFieldId
                    
                    if myField.CustomFieldDataType.CustomFieldDataTypeId = FIELDTYPE_LONGTEXT then
                        longValues.Add myValue
                    else
                        shortValues.Add myValue
                    end if
                next

                dim colsUsed
                for shortValueIndex = 0 to shortValues.Count - 1
                    if shortValueIndex = 0 then
                        response.Write "<tr class=""expand-child"">"
                        colsUsed = 0
                    elseif shortValueIndex mod 3 = 0 then
                        if COLUMNS_IN_ROW - colsUsed > 0 then response.Write "<td colspan=""" & COLUMNS_IN_ROW - colsUsed & """/>"
                        response.Write "</tr>"
                        response.Write "<tr class=""expand-child"">"
                        colsUsed = 0
                    end if

                    dim ShortField : set ShortField = new RegistrationCustomField
                    ShortField.LoadById shortValues(shortValueIndex).RegistrationCustomFieldId

                    response.Write "<td class=""heading"">" & ShortField.Name &"</td>"
        
                    if ShortField.CustomFieldDataType.CustomFieldDataTypeId = FIELDTYPE_OPTIONS then
                        dim optionValue : set optionValue = new RegistrationCustomOption
                        optionValue.RegistrationCustomFieldId = ShortField.RegistrationCustomFieldId
                        optionValue.Sequence = shortValues(shortValueIndex).Value
                        optionValue.Load

                        Response.Write "<td colspan=""2"">" & optionValue.Text & "</td>"
                    else
                        Response.Write "<td colspan=""2"">" & shortValues(shortValueIndex).Value & "</td>"
                    end if

                    colsUsed = colsUsed + 3
                next

                if shortValues.Count > 0 then 
                    if COLUMNS_IN_ROW - colsUsed > 0 then response.Write "<td colspan=""" & COLUMNS_IN_ROW - colsUsed & """/>"
                    response.Write "</tr>"
                end if

                for longValueIndex = 0 to longValues.Count - 1
                    dim longField : set longField = new RegistrationCustomField
                    longField.LoadById longValues(longValueIndex).RegistrationCustomFieldId

                    response.Write "<tr class=""expand-child"">"
                    response.Write "<td class=""heading"" nowrap=""nowrap"">" & longField.Name & "</td>"
                    Response.Write "<td colspan=""" & COLUMNS_IN_ROW - 1 & """>" & longValues(longValueIndex).Value & "</td>"
                    response.Write "</tr>"
                next
            end if                

            response.Write "<tr class=""expand-child"">"
            response.Write "<td class=""heading"" nowrap=""nowrap"">User Notes</td>"
            Response.Write "<td colspan=""" & COLUMNS_IN_ROW - 1 & """>" & myEvent.UserComments & "</td>"
            response.Write "</tr>"

            response.Write "<tr class=""expand-child"">"
            response.Write "<td class=""heading"" nowrap=""nowrap"">Admin Notes</td>"
            Response.Write "<td colspan=""" & COLUMNS_IN_ROW - 1 & """>" & myEvent.AdminComments & "</td>"
            response.Write "</tr>"

            response.Write "<tr class=""expand-child"">"
            Response.Write "<td colspan=""" & COLUMNS_IN_ROW & """ style=""background-color: grey"">&nbsp;</td>"
            response.Write "</tr>"
        next

        response.Write "</tbody>"
        
        '****************************************************
        ' footer
        '****************************************************
        Response.Write "<tfoot>"
        dim numberTitleColWritten : numberTitleColWritten = false

        for each statusKey in statusDict
            dim totalNumbers
            totalNumbers = split(statusDict(statusKey), ";")

            response.Write "<tr>"
            Response.Write "<td class=""heading"">" & statusKey & " Totals</td>"
            if numberTitleColWritten = false then
                Response.Write "<td colspan=""2"">" & totalNumbers(0) & "</td>"
                Response.Write "<td class=""heading"">Number of guests</td>"
                numberTitleColWritten = true
            else
                Response.Write "<td colspan=""3"">" & totalNumbers(0) & "</td>"
            end if
            Response.Write "<td colspan=""6"">" & totalNumbers(1) & "</td>"
            response.Write "</tr>"
        next

        response.Write "<tr>"
        Response.Write "<td class=""heading"">Totals</td>"
        if numberTitleColWritten = false then
            Response.Write "<td colspan=""2"">" & events.Count & "</td>"
            Response.Write "<td class=""heading"">Number of guests</td>"
            numberTitleColWritten = true
        else
            Response.Write "<td colspan=""3"">" & events.Count & "</td>"
        end if
        if usesDeposits then
            Response.Write "<td colspan=""6"">" & events.TotalPatrons & "</td>"
        else
            Response.Write "<td colspan=""5"">" & events.TotalPatrons & "</td>"
        end if
        response.Write "</tr>"
        Response.Write "</tfoot>"
        
        response.Write "</table>"
    end sub
%>