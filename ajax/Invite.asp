<!--#include file="../classes/IncludeList2.asp"-->
<%
    '*****************************************************************
    ' check the querystring to determine what action to take
    '*****************************************************************
    if Request.QueryString(QUERYSTRING_VAR_ACTION) = 0 then
        ' create a new invite
        SaveInvite Request.QueryString(QUERYSTRING_VAR_ID), _
                    Request.QueryString(QUERYSTRING_VAR_NAME), _
                    Request.QueryString(QUERYSTRING_VAR_EMAIL)
    elseif Request.QueryString(QUERYSTRING_VAR_ACTION) = 1 then
            SaveInviteFull DecodeId(Request.QueryString(QUERYSTRING_VAR_ID)), _
                            Request.QueryString(QUERYSTRING_VAR_NAME), _
                            Request.QueryString(QUERYSTRING_VAR_EMAIL), _
                            Request.QueryString(QUERYSTRING_VAR_STATUS), _
                            Request.QueryString(QUERYSTRING_VAR_NUMBEROFPATRONS)
    elseif Request.QueryString(QUERYSTRING_VAR_ACTION) = 5 then
        ' delete an invite
        DeleteInvite DecodeId(Request.QueryString(QUERYSTRING_VAR_ID))
    elseif Request.QueryString(QUERYSTRING_VAR_ACTION) = 6 then
        ' delete all invites for the event
        DeleteAllInvitesByEvent Request.QueryString(QUERYSTRING_VAR_ID)
    elseif Request.QueryString(QUERYSTRING_VAR_ACTION) = 10 then
        ' update an existing invite
        UpdateInvite DecodeId(Request.QueryString(QUERYSTRING_VAR_ID)), _
                     Request.QueryString(QUERYSTRING_VAR_NAME), _
                     Request.QueryString(QUERYSTRING_VAR_EMAIL)
    elseif request.QueryString(QUERYSTRING_VAR_ACTION) = 11 then
        ' update invite from full form
        UpdateInviteFull DecodeId(Request.QueryString(QUERYSTRING_VAR_ID)), _
                         Request.QueryString(QUERYSTRING_VAR_NAME), _
                         Request.QueryString(QUERYSTRING_VAR_EMAIL), _
                         Request.QueryString(QUERYSTRING_VAR_STATUS), _
                         Request.QueryString(QUERYSTRING_VAR_NUMBEROFPATRONS), _
                         Request.QueryString(QUERYSTRING_VAR_USERCOMMENTS)
    elseif Request.QueryString(QUERYSTRING_VAR_ACTION) = 12 then
        ' rsvp for an event
        RsvpInvite Request.QueryString(QUERYSTRING_VAR_ID), _
                   Request.QueryString(QUERYSTRING_VAR_NUMBEROFPATRONS), _
                   Request.QueryString(QUERYSTRING_VAR_STATUS), _
                   Request.QueryString(QUERYSTRING_VAR_USERCOMMENTS)
    elseif Request.QueryString(QUERYSTRING_VAR_ACTION) = 20 then
        ' build simple list
        BuildSimpleTable Request.QueryString(QUERYSTRING_VAR_ID)
    elseif Request.QueryString(QUERYSTRING_VAR_ACTION) = 21 then
        ' build simple list with editable entry
        BuildSimpleEditTable Request.QueryString(QUERYSTRING_VAR_ID), DecodeId(Request.QueryString(QUERYSTRING_VAR_NAME))
    elseif Request.QueryString(QUERYSTRING_VAR_ACTION) = 25 then
        ' build the full invite table form
        BuildFullTableHeading
        BuildFullTable DecodeId(Request.QueryString(QUERYSTRING_VAR_ID))
        BuildFullTableFooter
    elseif Request.QueryString(QUERYSTRING_VAR_aCTION) = 26 then
        ' build the full table with an edit row
        BuildFullTableHeading
        BuildFullEditTable DecodeId(Request.QueryString(QUERYSTRING_VAR_ID)), DecodeId(Request.QueryString(QUERYSTRING_VAR_NAME))
        BuildFullTableFooter
    elseif Request.QueryString(QUERYSTRING_VAR_ACTION) = 30 then
        ' send all of the invites for a new event
        SendAllNewInvites request.QueryString(QUERYSTRING_VAR_ID)
    elseif Request.QueryString(QUERYSTRING_VAR_ACTION) = 31 then
        ' send all of the invites that have not been sent
        SendAllUnsentInvites DecodeId(request.QueryString(QUERYSTRING_VAR_ID))
    elseif Request.QueryString(QUERYSTRING_VAR_ACTION) = 32 then
        ' send only the selected invite
        SendInvite DecodeId(request.QueryString(QUERYSTRING_VAR_ID))
    end if

    private sub BuildFullEditTable(EventId, InvitationId)
        dim eventInvites
        set eventInvites = new InvitationCollection

        eventInvites.LoadByEvent EventId

        with Response
            .Write "<table style=""width: auto; margin: 0px auto"">"
            .Write "<tr>"
            .Write "<th>Name</th>"
            .Write "<th>Email</th>"
            .Write "<th>Status</th>"
            .Write "<th>Count</th>"
            .Write "<th colspan=""3""/>"
            .Write "</tr>"

            for i = 0 to eventInvites.Count - 1
                if eventInvites(i).InvitationId = InvitationId then
                    .Write "<tr>"
                    .Write "<td>"
                    .Write "<input type=""hidden"" class=""InvitationId"" value=""" & eventInvites(i).InvitationId & """/>"
                    .Write "<input type=""text"" name=""txtGuestName"" id=""txtGuestName"" maxlength=""50"" value=""" & eventInvites(i).Name & """ /><br />"
                    .Write "<span name=""ValName"" style=""display: none; color: Red"">Name is required</span>"
                    .Write "</td>"
                    .Write "<td>"
                    .Write "<input type=""text"" name=""txtGuestEmail"" id=""txtGuestEmail"" maxlength=""120"" value=""" & eventInvites(i).EmailAddress & """ /><br />"
                    .Write "<span name=""ValEmail"" style=""display: none; color: Red"">Invalid email address</span>"
                    .Write "<span name=""ValEmailRequired"" style=""display: none; color: Red"">Email address required</span>"
                    .Write "</td>"
                    .Write "<td>"
                    .Write "<select name=""ddlRsvpStatus"" id=""ddlRsvpStatus"">"
                    .Write "<option " & iif(len(eventInvites(i).RsvpStatus) = 0, "selected", "") & "/>"
                    .Write "<option value=""Going"" " & iif(eventInvites(i).RsvpStatus = "Going", "selected", "") & ">Going</option>"
                    .Write "<option value=""Maybe"" " & iif(eventInvites(i).RsvpStatus = "Maybe", "selected", "") & ">Maybe</option>"
                    .Write "<option value=""Not Going"" " & iif(eventInvites(i).RsvpStatus = "Not Going", "selected", "") & ">Not Going</option>"
                    .Write "</select>"
                    .Write "</td>"
                    .Write "<td>"
                    .Write "<input type=""text"" name=""txtRsvpCount"" id=""txtRsvpCount"" value=""" & iif(eventInvites(i).RsvpCount > 0, eventInvites(i).RsvpCount, "") & """ style=""width: 50px""/><br/>"
                    .Write "<span name=""ValRsvpCount"" style=""display: none; color: Red"">Invalid number</span>"
                    .Write "</td>"
                    .Write "<td/>"
                    .Write "<td><a href=""javascript: updateInvite(" & EncodeId(eventInvites(i).InvitationId) & ");"">Update</a></td>"
                    .Write "<td><a href=""javascript: loadTable();"">Cancel</a></td>"
                    .Write "</tr>"

                    .Write "<tr>"
                    .Write "<td>Comments:</td>"
                    .Write "<td colspan=""6"" style=""text-align: right""><span name=""charcounter"">(" & (1000 - len(eventInvites(i).Comments)) & " characters available)</span></td>"
                    .Write "</tr>"
                    .Write "<tr>"
                    .Write "<td colspan=""7"">"
                    .Write "<textarea id=""UserComments"" style=""width: 100%"" maxlength=""1000"" onkeyup=""handleMaxLengthInvitation()"">" & eventInvites(i).Comments & "</textarea>"
                    .Write "</td>"
                    .Write "</tr>"
                else
                    .Write "<tr>"
                    .Write "<td><input type=""hidden"" class=""InvitationId"" value=""" & eventInvites(i).InvitationId & """/>" & eventInvites(i).Name & "</td>"
                    .Write "<td>" & eventInvites(i).EmailAddress & "</td>"
                    .Write "<td>"
                    if len(eventInvites(i).RsvpStatus) = 0 then
                        .Write "Unknown"
                    else
                        .Write eventInvites(i).RsvpStatus
                    end if
                    .Write "</td>"
                    .Write "<td>" & iif(eventInvites(i).RsvpCount > 0, eventInvites(i).RsvpCount, "") & "</td>"
                    if len(eventInvites(i).RsvpStatus) = 0 then
                        .Write "<td><a href=""javascript: resendInvite(" & EncodeId(eventInvites(i).InvitationId) & ");"">Resend</a></td>"
                    else
                        .Write "<td/>"
                    end if
                    .Write "<td><a href=""javascript: editInvite(" & EncodeId(eventInvites(i).InvitationId) & ");"">Edit</a></td>"
                    .Write "<td><a href=""javascript: deleteInvite(" & EncodeId(eventInvites(i).InvitationId) & ");"">Delete</a></td>"
                    .Write "</tr>"

                    if len(eventInvites(i).Comments) > 0 then
                        .Write "<tr>"
                        .Write "<td>Comments:</td>"
                        .Write "<td colspan=""6"">" & eventInvites(i).Comments & "</td>"
                        .Write "</tr>"
                    end if
                end if
            next

            .Write "</table>"
        end with
    end sub

    private sub BuildFullTable(EventId)
        dim eventInvites
        set eventInvites = new InvitationCollection

        eventInvites.LoadByEvent EventId

        with Response
            .Write "<table style=""width: auto; margin: 0px auto"">"
            .Write "<tr>"
            .Write "<th>Name</th>"
            .Write "<th>Email</th>"
            .Write "<th>Status</th>"
            .Write "<th>Count</th>"
            .Write "<th colspan=""3""/>"
            .Write "</tr>"

            for i = 0 to eventInvites.Count - 1
                .Write "<tr>"
                .Write "<td><input type=""hidden"" class=""InvitationId"" value=""" & eventInvites(i).InvitationId & """/>" & eventInvites(i).Name & "</td>"
                .Write "<td>" & eventInvites(i).EmailAddress & "</td>"
                .Write "<td>"
                if len(eventInvites(i).RsvpStatus) = 0 then
                    .Write "Unknown"
                else
                    .Write eventInvites(i).RsvpStatus
                end if
                .Write "</td>"
                .Write "<td>" & iif(eventInvites(i).RsvpCount > 0, eventInvites(i).RsvpCount, "") & "</td>"
                if len(eventInvites(i).RsvpStatus) = 0 then
                    .Write "<td><a href=""javascript: resendInvite(" & EncodeId(eventInvites(i).InvitationId) & ");"">Resend</a></td>"
                else
                    .Write "<td/>"
                end if
                .Write "<td><a href=""javascript: editInvite(" & EncodeId(eventInvites(i).InvitationId) & ");"">Edit</a></td>"
                .Write "<td><a href=""javascript: deleteInvite(" & EncodeId(eventInvites(i).InvitationId) & ");"">Delete</a></td>"
                .Write "</tr>"

                if len(eventInvites(i).Comments) > 0 then
                    .Write "<tr>"
                    .Write "<td>Comments:</td>"
                    .Write "<td colspan=""7"">" & eventInvites(i).Comments & "</td>"
                    .Write "</tr>"
                end if
            next

            .Write "<tr>"
            .Write "<td><input type=""text"" name=""txtGuestName"" id=""txtGuestName"" maxlength=""50"" /><br />"
            .Write "<span name=""ValName"" style=""display: none; color: Red"">Name is required</span></td>"
            .Write "<td><input type=""text"" name=""txtGuestEmail"" id=""txtGuestEmail"" maxlength=""120"" /><br />"
            .Write "<span name=""ValEmail"" style=""display: none; color: Red"">Invalid email address</span>"
            .Write "<span name=""ValEmailRequired"" style=""display: none; color: Red"">Email address required</span></td>"
            .Write "<td>"
            .Write "<select name=""ddlRsvpStatus"" id=""ddlRsvpStatus"">"
            .Write "<option />"
            .Write "<option value=""Going"">Going</option>"
            .Write "<option value=""Maybe"">Maybe</option>"
            .Write "<option value=""Not Going"">Not Going</option>"
            .Write "</select>"
            .Write "</td>"
            .Write "<td>"
            .Write "<input type=""text"" name=""txtRsvpCount"" id=""txtRsvpCount"" style=""width: 50px""/><br/>"
            .Write "<span name=""ValRsvpCount"" style=""display: none; color: Red"">Invalid number</span>"
            .Write "</td>"
            .Write "<td/>"
            .Write "<td><a href=""javascript: addInvite('true');"">Add</a></td>"
            .Write "<td/>"
            .Write "</tr>"
            .Write "</table>"
        end with
    end sub

    private sub BuildFullTableFooter()
        with response
            .Write "<p style=""text-align: center"">"
            .Write "<button onclick=""sendInvites()"" style=""text-align: center"">Send Invites</button>"
            .Write "</p>"
        end with
    end sub

    private sub BuildFullTableHeading()
        with response
            '.Write "<table cellpadding=""2"" cellspacing=""3"" border=""0"" style=""margin: 0 auto;"">"
            '.Write "<tr>"
            '.Write "<td colspan=""2"">"
            .Write "<ul style=""list-style: none; padding: 0px; font-size: 16 px; font-weight: bold; text-align: center"">"
            .Write "<li style=""display: inline; padding: 8px; cursor: pointer"" onclick=""showRegistrationForm()"">Event Registration</li>"
            .Write "<li style=""display: inline; padding: 8px; background-color: #444; color: white"">Party Invitations</li>"
            .Write "</ul>"
            '.Write "</td>"
            '.Write "</tr>"
            '.Write "</table>"
        end with
    end sub

    private sub BuildSimpleEditTable(EventId, InvitationId)
        dim eventInvites
        set eventInvites = new InvitationCollection

        eventInvites.LoadByEvent EventId

        with Response
            .Write "<table style=""width: auto; margin: 0px auto"">"
            .Write "<tr>"
            .Write "<th>Name</th>"
            .Write "<th>Email</th>"
            .Write "<th colspan=""2""/>"
            .Write "</tr>"

            for i = 0 to eventInvites.Count - 1
                .Write "<tr>"

                if eventInvites(i).InvitationId = InvitationId then
                    .Write "<td><input type=""hidden"" class=""InvitationId"" value=""" & eventInvites(i).InvitationId & """/>"
                    .Write "<input type=""text"" name=""txtGuestName"" id=""txtGuestName"" maxlength=""50"" value=""" & eventInvites(i).Name & """ /><br />"
                    .Write "<span name=""ValName"" style=""display: none; color: Red"">Name is required</span></td>"
                    .Write "<td><input type=""text"" name=""txtGuestEmail"" id=""txtGuestEmail"" maxlength=""120"" value=""" & eventInvites(i).EmailAddress & """ /><br />"
                    .Write "<span name=""ValEmail"" style=""display: none; color: Red"">Invalid email address</span>"
                    .Write "<span name=""ValEmailRequired"" style=""display: none; color: Red"">Email address required</span></td>"
                    .Write "<td><a href=""javascript: updateInvite(" & EncodeId(eventInvites(i).InvitationId) & ");"">Update</a></td>"
                    .Write "<td><a href=""javascript: loadTable();"">Cancel</a></td>"
                else
                    .Write "<td><input type=""hidden"" class=""InvitationId"" value=""" & eventInvites(i).InvitationId & """/>""" & eventInvites(i).Name & "</td>"
                    .Write "<td>" & eventInvites(i).EmailAddress & "</td>"
                    .Write "<td><a href=""javascript: editInvite(" & EncodeId(eventInvites(i).InvitationId) & ");"">Edit</a></td>"
                    .Write "<td><a href=""javascript: deleteInvite(" & EncodeId(eventInvites(i).InvitationId) & ");"">Delete</a></td>"
                end if

                .Write "</tr>"
            next

            .Write "</table>"
        end with
    end sub

    private sub BuildSimpleTable(EventId)
        dim eventInvites
        set eventInvites = new InvitationCollection

        eventInvites.LoadByEvent EventId

        with Response
            .Write "<table style=""width: auto; margin: 0px auto"">"
            .Write "<tr>"
            .Write "<th>Name</th>"
            .Write "<th>Email</th>"
            .Write "<th colspan=""2""/>"
            .Write "</tr>"

            for i = 0 to eventInvites.Count - 1
                .Write "<tr>"
                .Write "<td><input type=""hidden"" class=""InvitationId"" value=""" & eventInvites(i).InvitationId & """/>" & eventInvites(i).Name & "</td>"
                .Write "<td>" & eventInvites(i).EmailAddress & "</td>"
                .Write "<td><a href=""javascript: editInvite(" & EncodeId(eventInvites(i).InvitationId) & ");"">Edit</a></td>"
                .Write "<td><a href=""javascript: deleteInvite(" & EncodeId(eventInvites(i).InvitationId) & ");"">Delete</a></td>"
                .Write "</tr>"
            next

            .Write "<tr>"
            .Write "<td><input type=""text"" name=""txtGuestName"" id=""txtGuestName"" maxlength=""50"" /><br />"
            .Write "<span name=""ValName"" style=""display: none; color: Red"">Name is required</span></td>"
            .Write "<td><input type=""text"" name=""txtGuestEmail"" id=""txtGuestEmail"" maxlength=""120"" /><br />"
            .Write "<span name=""ValEmail"" style=""display: none; color: Red"">Invalid email address</span>"
            .Write "<span name=""ValEmailRequired"" style=""display: none; color: Red"">Email address required</span></td>"
            .Write "<td/>"
            .Write "<td><a href=""javascript: addInvite('false');"">Add</a></td>"
            .Write "</tr>"
            .Write "</table>"
        end with
    end sub

    private sub DeleteAllInvitesByEvent(EventId)
        dim eventInvites
        set eventInvites = new InvitationCollection

        eventInvites.LoadByEvent EventId

        for i = 0 to eventInvites.Count - 1
            myInvite(i).Delete myInvite(i).InvitationId
        next
    end sub

    private sub DeleteInvite(InvitationId)
        dim myInvite
        set myInvite = new Invitation
        myInvite.Delete InvitationId
    end sub

    'private sub MarkInviteSent(InvitationId)
    '    dim myInvite
    '    set myInvite = new Invitation

    '    with myInvite
    '        .Load InvitationId

    '        .InviteDateTime = Now()

    '        .Save
    '    end with
    'end sub

    private sub SaveInvite(EventId, Name, EmailAddress)
        dim myInvite
        set myInvite = new Invitation

        with myInvite
            .InvitationId = 0
            .EventId = EventId
            .Name = Name
            .EmailAddress = EmailAddress
            .InviteDateTime = ""
            .RsvpCount = 0
            .RsvpDateTime = ""

            .Save

            Response.Write .InvitationId
        end with
    end sub

    private sub SaveInviteFull(EventId, Name, EmailAddress, RsvpStatus, RsvpCount)
        dim myInvite
        set myInvite = new Invitation

        with myInvite
            .InvitationId = 0
            .EventId = EventId
            .Name = Name
            .EmailAddress = EmailAddress
            .InviteDateTime = ""
            .RsvpStatus = RsvpStatus
            .RsvpCount = RsvpCount
            .RsvpDateTime = Now()

            .Save

            Response.Write .InvitationId
        end with
    end sub

    private sub SendAllNewInvites(EventId)
        dim events
        set events = new InvitationCollection
        events.LoadByEvent EventId

        for i = 0 to events.Count - 1
            SendInvite events(i).InvitationId
        next
    end sub

    private sub SendAllUnsentInvites(EventId)
        dim events
        set events = new InvitationCollection
        events.LoadUnsentByEvent EventId
    
        for i = 0 to events.Count - 1
            SendInvite events(i).InvitationId
        next
    end sub

    private sub SendInvite(InviteId)
        dim myInvite, myInviteEmail
        set myInvite = new Invitation
        set myInviteEmail = new PartyInviteEmail

        with myInvite
            .Load InviteId

            myInviteEmail.SendByObject myInvite
    
            .InviteDateTime = Now()

            .Save

            Response.Write "Sent to: " & .Name & vbCrLf
        end with
    end sub

    private sub RsvpInvite(InvitationId, NumberOfGuests, Status, Comments)
        dim myInvite
        set myInvite = new Invitation

        with myInvite
            .Load InvitationId

            .RsvpCount = NumberOfGuests
            .RsvpDateTime = Now()
            .RsvpStatus = Status
            .Comments = Comments

            .Save
        end with
    end sub

    private sub UpdateInvite(InvitationId, Name, EmailAddress)
        dim myInvite
        set myInvite = new Invitation

        with myInvite
            .Load InvitationId

            .Name = Name
            .EmailAddress = EmailAddress

            .Save
        end with
    end sub

    private sub UpdateInviteFull(InvitationId, Name, EmailAddress, RsvpStatus, RsvpCount, Comments)
        dim myInvite
        set myInvite = new Invitation

        with myInvite
            .Load InvitationId

            .Name = Name
            .EmailAddress = EmailAddress
            .RsvpStatus = RsvpStatus
            .RsvpCount = RsvpCount
            .RsvpDateTime = Now()
            .Comments = Comments

            .Save
        end with
    end sub
%>