<!--#include file="../classes/IncludeList2.asp"-->
<% 
    '*****************************************************************
    ' check the querystring to determine what action to take
    '*****************************************************************
    if Request.QueryString(QUERYSTRING_VAR_ACTION) = 0 then
        SaveNewEvent Request.QueryString(QUERYSTRING_VAR_SELECTEDDATE), _
                     Request.QueryString(QUERYSTRING_VAR_STARTTIME), _
                     Request.QueryString(QUERYSTRING_VAR_NUMBEROFPATRONS), _
                     Request.QueryString(QUERYSTRING_VAR_AGEOFPATRONS), _
                     Request.QueryString(QUERYSTRING_VAR_EMAIL), _
                     Request.QueryString(QUERYSTRING_VAR_PHONE), _
                     Request.QueryString(QUERYSTRING_VAR_NAME), _
                     Request.QueryString(QUERYSTRING_VAR_EVENTTYPE), _
                     Request.QueryString(QUERYSTRING_VAR_PARTYNAME), _
                     Request.QueryString(QUERYSTRING_VAR_USERCOMMENTS), _
                     Request.QueryString(QUERYSTRING_VAR_GUNSIZE)
    elseif Request.QueryString(QUERYSTRING_VAR_ACTION) = 1 then
        WriteEventTableForDay Request.QueryString(QUERYSTRING_VAR_SELECTEDDATE)
    elseif Request.QueryString(QUERYSTRING_VAR_ACTION) = 2 then
        WriteEventTableRemaining
    elseif Request.QueryString(QUERYSTRING_VAR_ACTION) = 3 then
        dim eventId
        eventId = Request.QueryString(QUERYSTRING_VAR_ID)
        
        if len(Request.QueryString(QUERYSTRING_VAR_DECODE)) > 0 then
            if request.QueryString(QUERYSTRING_VAR_DECODE) = "true" then
                eventId = DecodeId(Request.QueryString(QUERYSTRING_VAR_ID))
            end if
        end if

        SaveExistingEvent eventId, _
                          Request.QueryString(QUERYSTRING_VAR_SELECTEDDATE), _
                          Request.QueryString(QUERYSTRING_VAR_STARTTIME), _
                          Request.QueryString(QUERYSTRING_VAR_NUMBEROFPATRONS), _
                          Request.QueryString(QUERYSTRING_VAR_AGEOFPATRONS), _
                          Request.QueryString(QUERYSTRING_VAR_EMAIL), _
                          Request.QueryString(QUERYSTRING_VAR_PHONE), _
                          Request.QueryString(QUERYSTRING_VAR_NAME), _
                          Request.QueryString(QUERYSTRING_VAR_EVENTTYPE), _
                          Request.QueryString(QUERYSTRING_VAR_PARTYNAME), _
                          Request.QueryString(QUERYSTRING_VAR_USERCOMMENTS), _
                          Request.QueryString(QUERYSTRING_VAR_ADMINCOMMENTS), _
                          Request.QueryString(QUERYSTRING_VAR_RESENDEMAIL), _
                          Request.QueryString(QUERYSTRING_VAR_GUNSIZE)
    elseif Request.QueryString(QUERYSTRING_VAR_ACTION) = 4 then
        dim cancelEventId
        cancelEventId = Request.QueryString(QUERYSTRING_VAR_ID)
        
        if len(Request.QueryString(QUERYSTRING_VAR_DECODE)) > 0 then
            if request.QueryString(QUERYSTRING_VAR_DECODE) = "true" then
                cancelEventId = DecodeId(Request.QueryString(QUERYSTRING_VAR_ID))
            end if
        end if

        CancelExistingEvent cancelEventId, Request.QueryString(QUERYSTRING_VAR_RESENDEMAIL)
    elseif Request.QueryString(QUERYSTRING_VAR_ACTION) = 5 then
        ShowRegistrationValidationForm "", true
    elseif Request.QueryString(QUERYSTRING_VAR_ACTION) = 6 then
        if VerifyEmail(Request.QueryString(QUERYSTRING_VAR_ID), Request.QueryString(QUERYSTRING_VAR_EMAIL)) then
            ShowRegistrationForm Request.QueryString(QUERYSTRING_VAR_ID)
        else
            ShowRegistrationValidationForm Request.QueryString(QUERYSTRING_VAR_EMAIL), false
        end if
    elseif Request.QueryString(QUERYSTRING_VAR_ACTION) = 7 then
        ' used for the daily stacked report
        BuildDailyStackedReportDataArray Request.QueryString(QUERYSTRING_VAR_SELECTEDDATE)
    elseif Request.QueryString(QUERYSTRING_VAR_ACTION) = 8 then
        ' used in the daily stacked report
        BuildPartyList Request.QueryString(QUERYSTRING_VAR_SELECTEDDATE)
    elseif Request.QueryString(QUERYSTRING_VAR_ACTION) = 9 then
        CheckOutEvent Request.QueryString(QUERYSTRING_VAR_ID), request.QueryString(QUERYSTRING_VAR_STARTTIME)
    elseif Request.QueryString(QUERYSTRING_VAR_ACTION) = 10 then
        ShowRegistrationForm Request.QueryString(QUERYSTRING_VAR_ID)
    end if

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
            columnString = columnString & "{""id"":"""",""label"":""" 
            if len(SelectedEvents(i).PartyName) > 0 then
                columnString = columnString &  SelectedEvents(i).PartyName
            else
                columnString = columnString & SelectedEvents(i).ContactName
            end if

            columnString = columnString & """,""pattern"":"""",""type"":""number""}"

            if i < SelectedEvents.Count - 1 then
                columnString = columnString &  ","
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
                dim personCount

                dim startDateTime, currentDateTime, endDateTime
                startDateTime = SelectedEvents(i).SelectedDate & " " & GetMilitaryTime(SelectedEvents(i).StartTime)
                currentDateTime = myDate.SelectedDate & " " & currentTime
                
                if len(selectedEvents(i).CheckOutTime) > 0 then
                    endDateTime = SelectedEvents(i).SelectedDate & " " & GetMilitaryTime(SelectedEvents(i).CheckOutTime)
                else
                    endDateTime = SelectedEvents(i).SelectedDate & " " & GetMilitaryTime(SelectedEvents(i).EstimatedEndTime)
                end if

                'dim diffInMinutes
                'if len(SelectedEvents(i).CutOffTime) > 0 then
                    'diffInMinutes = DateDiff("n", startDateTime, currentDateTime)
                    'diffInMinutes = DateDiff("n", cutOffTime, currentDateTime)
                'else
                    'diffInMinutes = DateDiff("n", startDateTime, currentDateTime)
                'end if

                dim startDiff, endDiff
                startDiff = DateDiff("n", startDateTime, currentDateTime)
                endDiff = DateDiff("n", endDateTime, currentDateTime)

                'response.Write "DEBUG: EventId = " & SelectedEvents(i).EventId & "<br/>"
                'response.Write "DEBUG: currentDateTime = " & currentDateTime & "<br/>"
                'response.Write "DEBUG: endDateTime = " & endDateTime & "<br/>"
                'response.Write "DEBUG: endDiff = " & endDiff & "<br/>"
                'response.Write "DEBUG: startDateTime = " & startDateTime & "<br/>"
                'response.Write "DEBUG: startDiff = " & startDiff & "<br/>"

                'if diffInMinutes >= 0 and diffInMinutes <= 150 then
                if startDiff >= 0 and endDiff <= 0 and len(SelectedEvents(i).NumberOfPatrons) > 0 then
                    personCount = SelectedEvents(i).NumberOfPatrons
                else
                    personCount = 0
                end if

                rowString = rowString & "{""v"":" & personCount & ",""f"":null}"

                if i < SelectedEvents.Count - 1 then
                    rowString = rowString & ","
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
                response.Write "<li>"

                dim partyName
                if len(myEvents(i).PartyName) = 0 then
                    partyName = myEvents(i).ContactName
                else
                    partyName = myEvents(i).PartyName
                end if

                response.Write partyName

                if len(myEvents(i).GunSize) > 0 then
                    response.Write " (" & myEvents(i).GunSizeShort & ")"
                end if

                response.Write "<img src=""../content/images/clock.png"" alt="""" title=""Check Out"" onclick='checkOut(" & myEvents(i).EventId & ", """ & replace(partyName, "'", "") & """);' />"
                Response.Write "</li>"
            next

            response.Write "</ul>" & vbCrLf
        end if
    end sub

    '**********************************************************
    ' Build an option list for the event types
    '**********************************************************
    sub BuildEventTypeOptions()
        dim types
        set types = new EventTypeCollection

        types.LoadAllActive

        for i = 0 to types.Count - 1
            dim myType
            set myType = types(i)

            response.Write "<option value='" & myType.EventTypeId & "'>" & myType.Description & "</option>"
        next
    end sub

    '**********************************************************
    ' build an option list for available times
    '**********************************************************
    sub BuildTimeOptions(OpenTime, CloseTime)
        dim currentTime, endTime
        currentTime = OpenTime
        endTime = FormatDateTime(DateAdd("h", -2, GetTimeString(CloseTime)), vbShortTime)

        while currentTime <= endTime
            response.Write "<option value='" & currentTime & "'>" & GetTimeString(currentTime) & "</option>"
            currentTime = FormatDateTime(DateAdd("n", 30, GetTimeString(currentTime)), vbShortTime)
        wend
    end sub

    '**********************************************************
    ' cancel the specified event and send an email if requested
    '**********************************************************
    sub CancelExistingEvent(EventId, SendEmail)
        dim myEvent, myEmail
        set myEvent = new ScheduledEvent
        set myEmail = new CancellationEmail

        myEvent.Delete EventId
        if SendEmail then myEmail.SendById EventId
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
    ' update the existing event
    '**********************************************************
    sub SaveExistingEvent(EventId, ScheduledDate, StartTime, NumberOfPatrons, AgeOfPatrons, ContactEmail, ContactPhone, ContactName, EventTypeId, PartyName, UserComments, AdminComments, ResendEmail, GunSize)
        dim myEvent
        set myEvent = new ScheduledEvent
        myEvent.Load(EventId)

        with myEvent        
            .SelectedDate = ScheduledDate
            .StartTime = StartTime
            .NumberOfPatrons = NumberOfPatrons
            .AgeOfPatrons = AgeOfPatrons
            .ContactEmailAddress = ContactEmail
            .ContactPhone = ContactPhone
            .ContactName = ContactName
            .EventTypeId = EventTypeId
            .PartyName = PartyName
            .UserComments = UserComments
            .AdminComments = AdminComments
            .GunSize = GunSize
            
            if .EventId = 0 then .ConfirmationDateTime = Now

            myEvent.Save 

            if ResendEmail then
                dim myEmail
                set myEmail = new ConfirmationEmail
                myEmail.SendByObject myEvent
            end if
        end with
    end sub

    '**********************************************************
    ' insert a new event
    '**********************************************************
    sub SaveNewEvent(ScheduledDate, StartTime, NumberOfPatrons, AgeOfPatrons, ContactEmail, ContactPhone, ContactName, EventTypeId, PartyName, UserComments, GunSize)
        dim myEvent, myEmail
        set myEvent = new ScheduledEvent
        set myEmail = new ConfirmationEmail

        with myEvent        
            .EventId = 0
            .SelectedDate = ScheduledDate
            .StartTime = StartTime
            .NumberOfPatrons = NumberOfPatrons
            .AgeOfPatrons = AgeOfPatrons
            .ContactEmailAddress = ContactEmail
            .ContactPhone = ContactPhone
            .ContactName = ContactName
            .EventTypeId = EventTypeId
            .PartyName = PartyName
            .ReminderDateTime = ""
            .UserComments = UserComments
            .AdminComments = ""
            .GunSize = GunSize
            myEvent.Save 

            myEmail.SendByObject myEvent

            .SaveConfirmationDate .EventId, Now

            Response.Write .EventId
        end with
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

        with Response
            if myEvent.Active then
                .Write "<ul style=""list-style: none; padding: 0px; font-size: 16 px; font-weight: bold; text-align: center;"">"
                .Write "<li style=""display: inline; padding: 8px; background-color: #444; color: white"">Event Registration</li>"
                .Write "<li style=""display: inline; padding: 8px; cursor: pointer"" onclick=""showInviteForm()"">Party Invitations</li>"
                .Write "</ul>"
                
                .Write "<table cellpadding=""2"" cellspacing=""3"" border=""0"" style=""margin: 0 auto;"">"
                .Write "<tr>"
                .Write "<td colspan=""2"">"
                .Write "<p class=""center"">Please Completely Fill Out The Registration Form</p>"
                .Write "</td>"
                '.Write "<table cellpadding=""2"" cellspacing=""3"" border=""0"" style=""margin: 0 auto;"">"
                .Write "</tr>"
                .Write "<tr>"
                .Write "<td>Date:</td>"
                .Write "<td>"
                .Write "<input type=""hidden"" name=""AdminComments"" value=""" & myEvent.AdminComments & """ />"
                .Write "<input type=""text"" name=""Date"" value=""" & FormatDateTime(myDate.SelectedDate, vbShortDate) & """ disabled=""disabled"" style=""width: 75px"" />"
                .Write "<button name=""btnCalendar"" onclick=""""><img src=""content/images/calendar_view_month.png"" alt=""calendar"" /></button>"
                .Write "</td>"
                .Write "<td/>"
                .Write "</tr>"
                .Write "<tr>"
                .Write "<td>Your Name:</td>"
                .Write "<td>"
                .Write "<input type=""text"" name=""Name"" maxlength=""100"" value=""" & myEvent.ContactName & """ /><br />"
                .Write "<span name=""ValName"" style=""display: none; color: Red"">Name is required</span>"
                .Write "</td>"
                .Write "<td style=""color: Red"">*</td>"
                .Write "</tr>"
                .Write "<tr>"
                .Write "<td>Hours of Operation:</td>"
                .Write "<td colspan=""2"">"
                .Write "<span name=""openHours""></span>"
                .Write "</td>"
                .Write "</tr>"
                .Write "<tr>"
                .Write "<td>What Time Do you Want to Start?:</td>"
                .Write "<td>"
                .Write "<select name=""Time"">"
                    BuildTimeOptions myDate.StartTime, myDate.EndTime
                .Write "</select><br />"
                .Write "<span name=""ValTime"" style=""display: none; color: Red"">Start time is required</span>"
                .Write "</td>"
                .Write "<td style=""color: Red"">*</td>"
                .Write "</tr>"
                .Write "<tr>"
                .Write "<td>Approximate Number Of Guests:</td>"
                .Write "<td>"
                .Write "<select name=""NumberOfGuests"">"
                .Write "<option value=""""></option>"
                .Write "<option value=""2"">2</option>"
                .Write "<option value=""3"">3</option>"
                .Write "<option value=""4"">4</option>"
                .Write "<option value=""5"">5</option>"
                .Write "<option value=""6"">6</option>"
                .Write "<option value=""7"">7</option>"
                .Write "<option value=""8"">8</option>"
                .Write "<option value=""9"">9</option>"
                .Write "<option value=""10"">10</option>"
                .Write "<option value=""11"">11 - 15</option>"
                .Write "<option value=""15"">15 - 20</option>"
                .Write "<option value=""20"">20 - 25</option>"
                .Write "<option value=""25"">25 - 30</option>"
                .Write "<option value=""30"">30 - 40</option>"
                .Write "<option value=""40"">40 - 50</option>"
                .Write "<option value=""50"">50 - 75</option>"
                .Write "<option value=""75"">75 - 100</option>"
                .Write "</select>"
                .Write "</td>"
                .Write "<td/>"
                .Write "</tr>"
                .Write "<tr>"
                .Write "<td>Average Age Of Guests:</td>"
                .Write "<td>"
                .Write "<input type=""text"" name=""AgeOfGuests"" maxlength=""100"" value=""" & myEvent.AgeOfPatrons & """ />"
                .Write "</td>"
                .Write "<td/>"
                .Write "</tr>"
                .Write "<tr>"
                .Write "<td>Name For Your Party:</td>"
                .Write "<td>"
                .Write "<input type=""text"" name=""PartyName"" maxlength=""100"" value=""" & myEvent.PartyName & """ />"
                .Write "</td>"
                .Write "<td />"
                .Write "</tr>"
                .Write "<tr>"
                .Write "<td>Your Email Address:</td>"
                .Write "<td>"
                .Write "<input type=""text"" name=""Email"" maxlength=""128"" value=""" & myEvent.ContactEmailAddress & """ /><br />"
                .Write "<span name=""ValEmail"" style=""display: none; color: Red"">Invalid email address</span>"
                .Write "<span name=""ValEmailRequired"" style=""display: none; color: Red"">Email address required</span>"
                .Write "</td>"
                .Write "<td style=""color: Red"">*</td>"
                .Write "</tr>"
                .Write "<tr>"
                .Write "<td>Please Confirm Email:</td>"
                .Write "<td>"
                .Write "<input type=""text"" name=""ConfirmEmail"" maxlength=""128"" value=""" & myEvent.ContactEmailAddress & """ /><br />"
                .Write "<span name=""ValEmailMatch"" style=""display: none; color: Red"">Email address does not match</span>"
                .Write "</td>"
                .Write "<td/>"
                .Write "</tr>"
                .Write "<tr>"
                .Write "<td>Contact Phone Number:</td>"
                .Write "<td>"
                .Write "<input type=""text"" name=""Phone"" maxlength=""20"" value=""" & myEvent.ContactPhone & """ /><br />"
                .Write "<span name=""ValPhoneRequired"" style=""display: none; color: Red"">Phone number required</span>"
                .Write "<span name=""ValPhoneFormat"" style=""display: none; color: Red"">Invalid phone number</span>"
                .Write "</td>"
                .Write "<td style=""color: Red"">*</td>"
                .Write "</tr>"           
                .Write "<tr>"
                .Write "<td>Event Type:</td>"
                .Write "<td>"
                .Write "<select name=""EventTypeId"">"
                .Write "<option value=""""></option>"
                    BuildEventTypeOptions
                .Write "</select>"
                .Write "</td>"
                .Write "<td/>"
                .Write "</tr>"
                .Write "<tr>"
                .Write "<td>User Comments:</td>"
                .Write "<td colspan=""2"">"
                .Write "<span name=""userCharCounter"">(1000 characters available)</span>"
                .Write "</td>"
                .Write "</tr>"
                .Write "<tr>"
                .Write "<td colspan=""3"">"
                .Write "<textarea name=""UserComments"" style=""width: 100%"" maxlength=""1000"" onkeyup=""handleMaxLength()"">" & myEvent.UserComments & "</textarea>"
                .Write "</td>"
                .Write "</tr>"
                .Write "<tr>"
                .Write "<td colspan=""3"" style=""text-align: center"">"
                .Write "<input type=""button"" value=""Save"" onclick=""saveRegistration()"" />"
                .Write "&nbsp;"
                .Write "<input type=""button"" value=""Cancel"" onclick=""redirectToHome();"" />"
                .Write "</td>"
                .Write "</tr>"
                .Write "<tr>"
                .Write "<td colspan=""3"" style=""text-align: center"">"
                .Write "<input type=""button"" value=""Cancel This Reservation"" onclick=""cancelRegistration()"" />"
                .Write "</td>"
                .Write "</tr>"
                .Write "</table>"
                .Write "<script type=""text/javascript"">"
                .Write "initializeForm('" & myEvent.EventTypeId & "', '" & myEvent.NumberOfPatrons & "', '" & GetMilitaryTime(myEvent.StartTime) & "');"
                .Write "</script>"
            else
                .Write "<h1 class=""center"">Event Cancelled</h1>"
                .Write "<p class=""center"">This event has been cancelled. You can <a href=""eventcalendar.asp"">book another event.</a></p>"
                .Write "<p class=""center""><a href=""http://www.gatsplat.com"">Return home</a></p>"
            end if
        end with
    end sub

    '**********************************************************
    ' build the registration validation form
    '**********************************************************
    sub ShowRegistrationValidationForm(EmailAddress, FirstAttempt)
        with Response
            .Write "<h1 class=""center"">Registration Verification</h1>"
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
            .Write "<td style=""color: Red"">*</td>"
            .Write "</tr>"
            .Write "<tr>"
            .Write "<td colspan=""3"" class=""center"">"
            .Write "<button name=""btnSubmit"" onclick=""submitVerification()"">Submit</button>"
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
        dim events, types, myDict
        set events = new ScheduledEventCollection
        set types = new EventTypeCollection
        events.LoadByDate ScheduledDate
        types.LoadAll

        set myDict = types.ToDictionary()

        '*********************************************************
        ' this format does not match the format from the 
        ' next method. cannot refactor to reuse code
        '*********************************************************
        response.Write "<table name=""scheduleTable"" cellpadding=""2"" cellspacing=""0"" border=""1"" style=""margin: 0 auto"">"
        response.Write "<thead>"
        response.Write "<tr>"
        'response.Write "<td class=""heading"">Event Id</td>"
        response.Write "<th class=""heading header"" nowrap=""nowrap"">Party Name</th>"
        response.Write "<th class=""heading"">Contact</th>"
        response.Write "<th class=""heading"">Email</th>"
        response.Write "<th class=""heading"">Phone</th>"
        response.Write "<th class=""heading"">Event Type</th>"
        response.Write "<th class=""heading"">Number</th>"
        response.Write "<th class=""heading"">Age</th>"
        response.Write "<th class=""heading"">Start</th>"
        response.Write "<th class=""heading"">Created</th>"
        response.Write "</tr>"
        response.Write "</thead>"
        response.Write "<tbody>"

        for i = 0 to events.Count - 1
            dim myEvent
            set myEvent = events(i)

            response.Write "<tr>"
            'response.Write "<td>" & myEvent.EventId & "</td>"
            response.Write "<td><a href=""#"" onclick=""editRegistration(" & myEvent.EventId & ")""><img src=""content/images/pencil.png"" alt="""" title=""Edit"" /></a> " & myEvent.PartyName & "</td>"
            response.Write "<td>" & myEvent.ContactName & "</td>"
            response.Write "<td>" & myEvent.ContactEmailAddress & "</td>"
            response.Write "<td>" & myEvent.ContactPhone & "</td>"
            response.Write "<td>" 
            if len(myEvent.EventTypeId) > 0 then response.Write myDict.Item(myEvent.EventTypeId) 
            response.Write "</td>"
            response.Write "<td>" & myEvent.NumberOfPatrons & "</td>"
            response.Write "<td>" & myEvent.AgeOfPatrons & "</td>"
            response.Write "<td>" & myEvent.StartTimeLongFormat & "</td>"
            response.Write "<td>" & myEvent.ConfirmationDateTime & "</td>"
            response.Write "</tr>"
            
            response.Write "<tr class=""expand-child"">"
            response.Write "<td class=""heading"" nowrap=""nowrap"">User Notes</td>"
            Response.Write "<td colspan=""8"">" & myEvent.UserComments & "</td>"
            response.Write "</tr>"

            response.Write "<tr class=""expand-child"">"
            response.Write "<td class=""heading"" nowrap=""nowrap"">Admin Notes</td>"
            Response.Write "<td colspan=""8"">" & myEvent.AdminComments & "</td>"
            response.Write "</tr>"

            response.Write "<tr class=""expand-child"">"
            Response.Write "<td colspan=""9"" style=""background-color: grey"">&nbsp;</td>"
            response.Write "</tr>"
        next

        response.Write "</tbody>"
        
        Response.Write "<tfoot>"
        Response.Write "<td class=""heading"">Totals</td>"
        Response.Write "<td>" & events.Count & "</td>"
        Response.Write "<td colspan=""2""/>"
        Response.Write "<td class=""heading"">Number of guests</td>"
        Response.Write "<td>" & events.TotalPatrons & "</td>"
        Response.Write "<td colspan=""3""/>"
        Response.Write "<tfoot>"
        
        response.Write "</table>"
    end sub

    '**********************************************************
    ' build the event table for all events from today and forward
    '**********************************************************
    sub WriteEventTableRemaining()
        dim events, types, myDict
        set events = new ScheduledEventCollection
        set types = new EventTypeCollection
        events.LoadRemaining
        types.LoadAll

        set myDict = types.ToDictionary()

        '*********************************************************
        ' this format does not match the format from the 
        ' previous method. cannot refactor to reuse code
        '*********************************************************
        response.Write "<table name=""scheduleTable"" cellpadding=""2"" cellspacing=""0"" border=""1"" style=""margin: 0 auto"">"
        response.Write "<thead>"
        response.Write "<tr>"
        'response.Write "<td class=""heading"">Event Id</td>"
        response.Write "<th class=""heading"" nowrap=""nowrap"">Party Name</th>"
        response.Write "<th class=""heading"">Contact</th>"
        response.Write "<th class=""heading"">Email</th>"
        response.Write "<th class=""heading"">Phone</th>"
        response.Write "<th class=""heading"">Event Type</th>"
        response.Write "<th class=""heading"">Number</th>"
        response.Write "<th class=""heading"">Age</th>"
        response.Write "<th class=""heading"">Scheduled</th>"
        response.Write "<th class=""heading"">Start</th>"
        response.Write "<th class=""heading"">Created</th>"
        response.Write "</tr>"
        response.Write "</thead>"
        response.Write "<tbody>"

        for i = 0 to events.Count - 1
            dim myEvent
            set myEvent = events(i)

            response.Write "<tr>"
            'response.Write "<td>" & myEvent.EventId & "</td>"
            response.Write "<td><a href=""#"" onclick=""editRegistration(" & myEvent.EventId & ")""><img src=""content/images/pencil.png"" alt="""" title=""Edit"" /></a> " & myEvent.PartyName & "</td>"
            response.Write "<td>" & myEvent.ContactName & "</td>"
            response.Write "<td>" & myEvent.ContactEmailAddress & "</td>"
            response.Write "<td>" & myEvent.ContactPhone & "</td>"
            response.Write "<td>" 
            if len(myEvent.EventTypeId) > 0 then response.Write myDict.Item(myEvent.EventTypeId) 
            response.Write "</td>"
            response.Write "<td>" & myEvent.NumberOfPatrons & "</td>"
            response.Write "<td>" & myEvent.AgeOfPatrons & "</td>"
            response.Write "<td>" & myEvent.SelectedDate & "</td>"
            response.Write "<td>" & myEvent.StartTimeLongFormat & "</td>"
            response.Write "<td>" & myEvent.ConfirmationDateTime & "</td>"
            response.Write "</tr>"
            
            response.Write "<tr class=""expand-child"">"
            response.Write "<td class=""heading"" nowrap=""nowrap"">User Notes</td>"
            Response.Write "<td colspan=""9"">" & myEvent.UserComments & "</td>"
            response.Write "</tr>"

            response.Write "<tr class=""expand-child"">"
            response.Write "<td class=""heading"" nowrap=""nowrap"">Admin Notes</td>"
            Response.Write "<td colspan=""9"">" & myEvent.AdminComments & "</td>"
            response.Write "</tr>"

            response.Write "<tr class=""expand-child"">"
            Response.Write "<td colspan=""10"" style=""background-color: grey"">&nbsp;</td>"
            response.Write "</tr>"
        next

        response.Write "</tbody>"
        
        Response.Write "<tfoot>"
        Response.Write "<td class=""heading"">Totals</td>"
        Response.Write "<td>" & events.Count & "</td>"
        Response.Write "<td colspan=""2""/>"
        Response.Write "<td class=""heading"">Number of guests</td>"
        Response.Write "<td>" & events.TotalPatrons & "</td>"
        Response.Write "<td colspan=""4""/>"
        Response.Write "<tfoot>"
        
        response.Write "</table>"
    end sub
%>