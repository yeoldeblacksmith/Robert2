<!--#include file="../classes/IncludeList.asp"-->
<% 
    dim myDate
    set myDate = new AvailableDate

    if Request.QueryString("av") = 0 then
        myDate.Delete Request.QueryString("dt")
    elseif request.QueryString("av") = 1 then
        myDate.Save Request.QueryString("dt"), Request.QueryString("st"), Request.QueryString("et")
    elseif Request.QueryString("av") = 2 then
        BuildAvailableTimeOptions Request.QueryString(QUERYSTRING_VAR_SELECTEDDATE)
    elseif Request.QueryString("av") = 3 then
        BuildHoursOfOperation Request.QueryString(QUERYSTRING_VAR_SELECTEDDATE)
    end if

    sub BuildAvailableTimeOptions(SelectedDate)
        myDate.Load SelectedDate

        dim currentTime, endTime
        currentTime = myDate.StartTime
        endTime = FormatDateTime(DateAdd("h", -2, myDate.EndTimeLongFormat), vbShortTime)

        while currentTime <= endTime
            response.Write "<option value='" & currentTime & "'>" & GetTimeString(currentTime) & "</option>"
            'currentTime = FormatDateTime(DateAdd("n", 30, GetTimeString(currentTime)), vbShortTime)
            currentTime = FormatDateTime(DateAdd("n", CInt(Settings(SETTING_BOOKING_INTERVAL)), GetTimeString(currentTime)), vbShortTime)
        wend
    end sub

    sub BuildHoursOfOperation(SelectedDate)
        myDate.Load SelectedDate

        response.Write myDate.StartTimeLongFormat & " - " & myDate.EndTimeLongFormat
    end sub
%>