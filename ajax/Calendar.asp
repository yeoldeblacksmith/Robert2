<!--#include file="../classes/IncludeList.asp"-->
<%

    const CALENDARTYPE_ADMIN = "ca"
    const CALENDARTYPE_MINI = "ci"
    const CALENDARTYPE_MOBILE = "cm"
    const CALENDARTYPE_USER = "cu"
    const CALENDARTYPE_OPEN = "op"
    
    if len(Request.QueryString(QUERYSTRING_VAR_CALENDARTYPE)) > 0 then
        dim selectedDate 
        selectedDate = request.QueryString(QUERYSTRING_VAR_SELECTEDDATE)

        select case Request.QueryString(QUERYSTRING_VAR_CALENDARTYPE)
            case CALENDARTYPE_ADMIN
                WriteAdminCalendar selectedDate
            case CALENDARTYPE_MINI
                WriteMiniCalendar selectedDate
            case CALENDARTYPE_MOBILE
                WriteMobileCalendar selectedDate
            case CALENDARTYPE_OPEN
                WriteOpenCalendar selectedDate
            case CALENDARTYPE_USER
                WriteUserCalendar selectedDate
        end select
    end if

    sub WriteAdminCalendar(dtDay)
        dim currentMonth, currentDay, currentYear, _
            beginingOfMonth, endOfMonth, currentDate, _
            writingDate

        currentMonth = Month(dtDay)
        currentDay = Day(dtDay)
        currentYear = Year(dtDay)
        today = CDate(Month(Now()) & "/" & Day(Now()) & "/" & Year(Now()))

        beginingOfMonth = CDate(currentMonth & "/1/" & currentYear)
        endOfMonth = DateAdd("d", -1, DateAdd("m", 1, beginingOfMonth))
        writingDate = false

        Response.Write "<table cellpadding=""0"" cellspacing=""0"" border=""1"" class=""calendar"">"
        WriteHeading

        for i = 0 to 5
            Response.Write "<tr class=""week"">"
                    
            for j = 1 to 7
                if i = 0 then
                    if j = Weekday(beginingOfMonth) then
                        currentDate = beginingOfMonth
                        writingDate = true
                    end if
                end if

                if writingDate then 
                    dim openCellTag, myDate, exists

                    if currentDate >= today then
                        openCellTag = "<td onclick='changeOpenHours(""" & currentDate & """)' class='mouseover "

                        'availableDate = GetAvailableDate(currentDate)
                        set myDate = new AvailableDate
                        exists = myDate.Exists(currentDate)

                        if exists then
                            openCellTag = openCellTag & "available"
                        end if
                        openCellTag = openCellTag & "'>"
                    else
                        openCellTag = "<td class='unavailable'>"
                    end if

                    response.Write openCellTag

                    response.Write "<div style='float: left; margin-left: 5px'>" & Day(currentDate) & "</div>"
                    if exists then
                        myDate.Load(currentDate)
                        response.Write "<div style='float: right; text-align: right; margin-right: 5px'>" & _
                                       myDate.StartTimeLongFormat & _
                                       "<br/>-<br/>" & _
                                       myDate.EndTimeLongFormat & _
                                       "</div>"
                    end if
                    response.Write "</td>"
                    currentDate = DateAdd("d", 1, currentDate)
                else
                    response.Write "<td class='unavailable'/>"
                end if

                if currentDate > endOfMonth then writingDate = false
            next

            Response.Write "</tr>"

            if writingDate = false then i = 10
        next

        Response.Write "</table>"
    end sub

    sub WriteMiniCalendar(dtDay)
        dim currentMonth, currentDay, currentYear, _
            beginingOfMonth, endOfMonth, currentDate, _
            today, writingDate

        currentMonth = Month(dtDay)
        currentDay = Day(dtDay)
        currentYear = Year(dtDay)
        today = CDate(Month(Now()) & "/" & Day(Now()) & "/" & Year(Now()))

        beginingOfMonth = CDate(currentMonth & "/1/" & currentYear)
        endOfMonth = DateAdd("d", -1, DateAdd("m", 1, beginingOfMonth))
        writingDate = false

        Response.Write "<table cellpadding=""0"" cellspacing=""0"" border=""1"" class=""calendarMini"">" & vbCrLf
        WriteHeading

        for i = 0 to 5
            Response.Write "<tr class=""weekMini"">" & vbCrLf
                    
            for j = 1 to 7
                if i = 0 then
                    if j = Weekday(beginingOfMonth) then
                        currentDate = beginingOfMonth
                        writingDate = true
                    end if
                end if

                if writingDate then 
                    dim openCellTag, myDate, exists

                    if currentDate >= today then
                        openCellTag = "<td"

                        set myDate = new AvailableDate
                        exists = myDate.Exists(currentDate)

                        if exists then
                            openCellTag = openCellTag & " onclick='setDate(""" & currentDate & """)' class='mouseover available'"
                        end if
                        openCellTag = openCellTag & ">" & vbCrLf
                    else
                        openCellTag = "<td class='unavailable'>" & vbCrLf
                    end if

                    response.Write openCellTag

                    response.Write "<div style='float: left; margin-left: 5px'>" & Day(currentDate) & "</div>" & vbCrLf
                    if exists then
                        'myDate.Load(currentDate)
                        'response.Write "<div style='float: right; text-align: right; margin-right: 5px'>" & _
                        '               myDate.StartTimeLongFormat & _
                        '               "<br/>-<br/>" & _
                        '               myDate.EndTimeLongFormat & _
                        '               "</div>" & vbCrLf
                    end if
                    response.Write "</td>" & vbCrLf
                    currentDate = DateAdd("d", 1, currentDate)
                else
                    response.Write "<td class='unavailable'/>" & vbCrLf
                end if

                if currentDate > endOfMonth then writingDate = false
            next

            Response.Write "</tr>" & vbCrLf

            if writingDate = false then i = 10
        next

        Response.Write "</table>" & vbCrLf
    end sub

    sub WriteMobileCalendar(dtDay)
        dim currentMonth, currentDay, currentYear, _
            beginingOfMonth, endOfMonth, currentDate, _
            today, writingDate

        currentMonth = Month(dtDay)
        currentDay = Day(dtDay)
        currentYear = Year(dtDay)
        today = CDate(Month(Now()) & "/" & Day(Now()) & "/" & Year(Now()))

        beginingOfMonth = CDate(currentMonth & "/1/" & currentYear)
        endOfMonth = DateAdd("d", -1, DateAdd("m", 1, beginingOfMonth))
        writingDate = false

        Response.Write "<table cellpadding=""0"" cellspacing=""0"" border=""1"" class=""calendar"">" & vbCrLf
        WriteHeading

        for i = 0 to 5
            Response.Write "<tr class=""weekMini"">" & vbCrLf
                    
            for j = 1 to 7
                if i = 0 then
                    if j = Weekday(beginingOfMonth) then
                        currentDate = beginingOfMonth
                        writingDate = true
                    end if
                end if

                if writingDate then 
                    dim openCellTag, myDate, exists

                    if currentDate >= today then
                        openCellTag = "<td"

                        set myDate = new AvailableDate
                        exists = myDate.Exists(currentDate)

                        if exists then
                            openCellTag = openCellTag & " onclick='setDate(""" & currentDate & """)' class='mouseover available'"
                        end if
                        'openCellTag = openCellTag & ">" & vbCrLf
                        openCellTag = openCellTag & ">"
                    else
                        'openCellTag = "<td class='unavailable'>" & vbCrLf
                        openCellTag = "<td class='unavailable'>"
                    end if

                    response.Write openCellTag

                    'response.Write "<div style='float: left; margin-left: 5px'>" & Day(currentDate) & "</div>" & vbCrLf
                    response.Write Day(currentDate)
                    if exists then
                        'myDate.Load(currentDate)
                        'response.Write "<div style='float: right; text-align: right; margin-right: 5px'>" & _
                        '               myDate.StartTimeLongFormat & _
                        '               "<br/>-<br/>" & _
                        '               myDate.EndTimeLongFormat & _
                        '               "</div>" & vbCrLf
                    end if
                    response.Write "</td>" & vbCrLf
                    currentDate = DateAdd("d", 1, currentDate)
                else
                    response.Write "<td class='unavailable'/>" & vbCrLf
                end if

                if currentDate > endOfMonth then writingDate = false
            next

            Response.Write "</tr>" & vbCrLf

            if writingDate = false then i = 10
        next

        Response.Write "</table>" & vbCrLf
    end sub

    '******************************************************************
    ' this is the same as the mini calendar just all days are selectable.
    ' this is used for waivers when the registration module is disabled.
    '******************************************************************
    sub WriteOpenCalendar(dtDay)
        dim currentMonth, currentDay, currentYear, _
            beginingOfMonth, endOfMonth, currentDate, _
            today, writingDate

        currentMonth = Month(dtDay)
        currentDay = Day(dtDay)
        currentYear = Year(dtDay)
        today = CDate(Month(Now()) & "/" & Day(Now()) & "/" & Year(Now()))

        beginingOfMonth = CDate(currentMonth & "/1/" & currentYear)
        endOfMonth = DateAdd("d", -1, DateAdd("m", 1, beginingOfMonth))
        writingDate = false

        Response.Write "<table cellpadding=""0"" cellspacing=""0"" border=""1"" class=""calendarMini"">" & vbCrLf
        WriteHeading

        for i = 0 to 5
            Response.Write "<tr class=""weekMini"">" & vbCrLf
                    
            for j = 1 to 7
                if i = 0 then
                    if j = Weekday(beginingOfMonth) then
                        currentDate = beginingOfMonth
                        writingDate = true
                    end if
                end if

                if writingDate then 
                    dim openCellTag, myDate, exists

                    if currentDate >= today then
                        openCellTag = "<td"
                        openCellTag = openCellTag & " onclick='setDate(""" & currentDate & """)' class='mouseover available'"
                        openCellTag = openCellTag & ">" & vbCrLf
                    else
                        openCellTag = "<td class='unavailable'>" & vbCrLf
                    end if

                    response.Write openCellTag

                    response.Write "<div style='float: left; margin-left: 5px'>" & Day(currentDate) & "</div>" & vbCrLf
                    response.Write "</td>" & vbCrLf
                    currentDate = DateAdd("d", 1, currentDate)
                else
                    response.Write "<td class='unavailable'/>" & vbCrLf
                end if

                if currentDate > endOfMonth then writingDate = false
            next

            Response.Write "</tr>" & vbCrLf

            if writingDate = false then i = 10
        next

        Response.Write "</table>" & vbCrLf
    end sub

    sub WriteUserCalendar(dtDay)
        dim currentMonth, currentDay, currentYear, _
            beginingOfMonth, endOfMonth, currentDate, _
            today, writingDate

        currentMonth = Month(dtDay)
        currentDay = Day(dtDay)
        currentYear = Year(dtDay)
        today = CDate(Month(Now()) & "/" & Day(Now()) & "/" & Year(Now()))

        beginingOfMonth = CDate(currentMonth & "/1/" & currentYear)
        endOfMonth = DateAdd("d", -1, DateAdd("m", 1, beginingOfMonth))
        writingDate = false

        Response.Write "<table cellpadding=""0"" cellspacing=""0"" border=""1"" class=""calendar"">"
        WriteHeading

        for i = 0 to 5
            Response.Write "<tr class=""week"">"
                    
            for j = 1 to 7
                if i = 0 then
                    if j = Weekday(beginingOfMonth) then
                        currentDate = beginingOfMonth
                        writingDate = true
                    end if
                end if

                if writingDate then 
                    dim openCellTag, myDate, exists

                    if currentDate >= today then
                        openCellTag = "<td"

                        set myDate = new AvailableDate
                        exists = myDate.Exists(currentDate)

                        if exists then
                            openCellTag = openCellTag & " onclick='showRegistration(""" & currentDate & """)' class=""mouseover available"""
                        end if
                        openCellTag = openCellTag & ">"
                    else
                        openCellTag = "<td class=""unavailable"">"
                    end if

                    response.Write openCellTag

                    response.Write "<div style='float: left; margin-left: 5px'>" & Day(currentDate) & "</div>"
                    if exists then
                        myDate.Load(currentDate)
                        response.Write "<div style='float: right; text-align: right; margin-right: 5px'>" & _
                                       myDate.StartTimeLongFormat & _
                                       "<br/>-<br/>" & _
                                       myDate.EndTimeLongFormat & _
                                       "</div>"
                    end if
                    response.Write "</td>"
                    currentDate = DateAdd("d", 1, currentDate)
                else
                    response.Write "<td class=""unavailable""/>"
                end if

                if currentDate > endOfMonth then writingDate = false
            next

            Response.Write "</tr>"

            if writingDate = false then i = 10
        next

        Response.Write "</table>"
    end sub

    sub WriteHeading()
        with response
            .Write("<tr>" & vbCrLf)
            .Write("<td class=""heading"">Sun</td>" & vbCrLf)
            .Write("<td class=""heading"">Mon</td>" & vbCrLf)
            .Write("<td class=""heading"">Tue</td>" & vbCrLf)
            .Write("<td class=""heading"">Wed</td>" & vbCrLf)
            .Write("<td class=""heading"">Thu</td>" & vbCrLf)
            .Write("<td class=""heading"">Fri</td>" & vbCrLf)
            .Write("<td class=""heading"">Sat</td>" & vbCrLf)
            .Write("</tr>" & vbCrLf)
        end with
    end sub
%>