<!--#include file="../classes/IncludeList.asp"-->
<%
    dim myEvent, bEditMode, bDebugMode

    bDebugMode = false

    if len(request.QueryString(QUERYSTRING_VAR_ID)) > 0 then
        set myEvent = new ScheduledEvent
        myEvent.Load DecodeId(Request.QueryString(QUERYSTRING_VAR_ID))

        bEditMode = cbool( cint(request.QueryString(QUERYSTRING_VAR_NUMBEROFPATRONS)) <= cint(myEvent.NumberOfPatrons)  )
    else
        bEditMode = false
    end if

    dim timeList
    set timeList = GetTimeList(request.QueryString(QUERYSTRING_VAR_SELECTEDDATE), _
                               request.QueryString(QUERYSTRING_VAR_NUMBEROFPATRONS))
%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" style="min-height: 0px !important">
<head>
    <title></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <meta name="viewport" content="width=device-width, maximum-scale=1.0, minimum-scale=1.0" /> 

    <link type="text/css" rel="stylesheet" href="../content/css/eventregistration.css" />
    <link type="text/css" rel="stylesheet" href="../content/css/waiver.css" />
    <link type="text/css" rel="stylesheet" href="../content/css/jquery.mobile-1.1.0.css" />

    <script type="text/javascript" src="../content/js/jquery-1.7.2.min.js"></script>
    <script type="text/javascript" src="../content/js/jquery.mobile-1.1.0.min.js"></script>
    <script type="text/javascript">
        $(document).ready(function () {
            //$.mobile.activePage.css('min-height', $("div.waiverSection").height());
            window.parent.resizeFancyBox();
        });

        function closeWithoutTime() {
            window.parent.hideAvailableTimeWindow();
        }

        function closeWithTime(value, text) {
            window.parent.setStartTime(value, text);
            window.parent.hideAvailableTimeWindow();
        }
    </script>
</head>
<body>
    <div class="waiverSection blockCenter">
<%  if timeList.Count = 0 then %>
        <p class="center">Sorry we are too busy on this day for your group. Please check another day, or you could try later in case we get some cancellations.</p>
        <button class="center" onclick="closeWithoutTime()">Close Window</button>
<%  else %>
        <p class="center">Note: If a time is not available below, it means that time is booked. Simply choose a time before or after that. Any shown times are available to start.</p>
        <p>Select a start time:</p>
        
        <table style="width: 300px; margin: 0 auto">
        <%
            if timeList.Count = 1 then
                Response.Write "<tr>" & vbCrLf
                Response.Write "<td width=""50px"">&nbsp;</td>" & vbCrLf
                Response.Write "<td width=""50px"">&nbsp;</td>" & vbCrLf
                Response.Write "<td>" & vbCrLf
                response.Write "<button onclick=""closeWithTime('" & timeList(0) & "', '" & GetTimeString(timeList(0)) & _
                               "')"" data-mini=""true"" data-theme=""b"">" & GetTimeString(timeList(0)) & "</button>" & vbCrLf
                Response.Write "</td>" & vbCrLf
                Response.Write "<td width=""50px"">&nbsp;</td>" & vbCrLf
                Response.Write "<td width=""50px"">&nbsp;</td>" & vbCrLf
                Response.Write "</tr>" & vbCrLf
            else
                dim index : index = 0
                
                while index < timeList.Count
                    Response.Write "<tr>" & vbCrLf
                    Response.Write "<td width=""50px"">&nbsp;</td>" & vbCrLf
                    Response.Write "<td>" & vbCrLf
                    response.Write "<button onclick=""closeWithTime('" & timeList(index) & "', '" & GetTimeString(timeList(index)) & _
                                   "')"" data-mini=""true"" data-theme=""b"">" & GetTimeString(timeList(index)) & "</button>" & vbCrLf
                    index = index + 1
                    Response.Write "</td>" & vbCrLf
                    Response.Write "<td width=""50px"">&nbsp;</td>" & vbCrLf
                    Response.Write "<td>" & vbCrLf
                    if index < timeList.Count then
                        response.Write "<button onclick=""closeWithTime('" & timeList(index) & "', '" & GetTimeString(timeList(index)) & _
                                       "')"" data-mini=""true"" data-theme=""b"">" & GetTimeString(timeList(index)) & "</button>" & vbCrLf
                        index = index + 1
                    end if
                    Response.Write "</td>" & vbCrLf
                    Response.Write "<td width=""50px"">&nbsp;</td>" & vbCrLf
                    Response.Write "</tr>" & vbCrLf
                wend
            end if
        %>
        </table>
<%  end if %>
    </div>
</body>
</html>

<%
    '****************************************************************
    ' apply the max limit filters and return the completed list
    '****************************************************************
    private function GetFinalList(EventDate, SoftLimitList, FullTimeDictionary, NumberOfGuests)
        dim finalTimeList : set finalTimeList = new ArrayList

        for each key in FullTimeDictionary
            if bDebugMode then Response.Write "*=*=*=*=*=*=*=*=<br/>"
            if bDebugMode then response.Write "DEBUG: time " & key & "<br/>"
            if SoftLimitList.contains(key) = false then
                dim playersInSlot : playersInSlot = cint(FullTimeDictionary(key))
                    
                dim validTimeKey : validTimeKey = true
                startDateTime = EventDate.SelectedDate & " " & key
                endDateTime = EventDate.SelectedDate & " " & FormatDateTime(DateAdd("n", cint(Settings(SETTING_REGISTRATION_PARTYDURATION)) - 1, key), vbShortTime)

                for each innerKey in FullTimeDictionary
                    currentDateTime = EventDate.SelectedDate & " " & innerKey
                    startDiff = DateDiff("n", startDateTime, currentDateTime)
                    endDiff = DateDiff("n", endDateTime, currentDateTime)
                    if bDebugMode then response.Write "DEBUG: innerKey " & innerKey & "<br/>"
                    if bDebugMode then response.Write "DEBUG: startDiff " & startDiff & "<br/>"
                    if bDebugMode then response.Write "DEBUG: endDiff " & endDiff & "<br/>"

                    dim attemptedPlayerCount : attemptedPlayerCount = cint(FullTimeDictionary(innerKey)) + cint(NumberOfGuests)
                    if bDebugMode then response.Write "DEBUG: attemptedPlayerCount " & attemptedPlayerCount & "<br/>"

                    if (startDiff >= 0 and endDiff <= 0) and attemptedPlayerCount > CInt(Settings(SETTING_REGISTRATION_PLAYERLIMIT)) then
                        validTimeKey = false
                    end if           
                    
                    if bDebugMode then response.Write "DEBUG: validTimeKey " & validTimeKey & "<br/>"         
                next

                if validTimeKey then
                    if bDebugMode then response.Write "DEBUG: writing time " & key & " (" & playersInSlot & ")<br/>"
                    finalTimeList.Add key
                end if
            end if
        next

        set GetFinalList = finalTimeList
    end function

    '****************************************************************
    ' build the dictionary of all possible times for the selected day
    '****************************************************************
    private function GetFullTimeDictionary(SelectedDate)
        dim currentTime, endTime
        currentTime = FormatDateTime(SelectedDate.StartTimeLongFormat, vbShortTime)
        endTime = FormatDateTime(DateAdd("h", -2, SelectedDate.EndTimeLongFormat), vbShortTime)

        ' create a dictionary of times with the value being for the count of users
        dim timeDictionary : set timeDictionary = Server.CreateObject("Scripting.Dictionary")

        while currentTime <= endTime
            timeDictionary.Add currentTime, 0
            currentTime = FormatDateTime(DateAdd("n", CInt(Settings(SETTING_BOOKING_INTERVAL)), GetTimeString(currentTime)), vbShortTime)
        wend

        set GetFullTimeDictionary = timeDictionary
    end function

    '****************************************************************
    ' build a list of all times that hit the soft limit
    '****************************************************************
    private function GetSoftLimitList(SelectedDate, FullDictionary, EventList)
        dim invalidTimeList : set invalidTimeList = new ArrayList

        '' remove all keys that have reached the soft limit
        ' look for entries that have reached the soft limit and put them in a list of invalid times
        for each timeSlotKey in FullDictionary
            dim playerCount : playerCount = 0

            for i =  0 to EventList.Count - 1
                if GetMilitaryTime(EventList(i).StartTime) = timeSlotKey then
                    if len(EventList(i).NumberOfPatrons) > 0 and EventList(i).CanCountPlayers then
                        playerCount = playerCount + EventList(i).NumberOfPatrons
                    end if
                end if
            next
            
            if bDebugMode then response.Write "DEBUG: time " & timeSlotKey & " (" & playerCount & ")<br/>"

            if playerCount >= cint(Settings(SETTING_REGISTRATION_MAXPLAYERS)) then
                if bEditMode then
                    if timeSlotKey = myEvent.StartTimeShortFormat then
                    else
                        if FullDictionary.Exists(timeSlotKey) then 
                            invalidTimeList.add timeSlotKey
                            if bDebugMode then response.Write "DEBUG: Removing time " & timeSlotKey & " (" & playerCount & ")<br/>"
                        end if
                    end if
                else
                    if FullDictionary.Exists(timeSlotKey) then 
                        invalidTimeList.add timeSlotKey
                        if bDebugMode then response.Write "DEBUG: Removing time " & timeSlotKey & " (" & playerCount & ")<br/>"
                    end if
                end if
            end if
        next

        set GetSoftLimitList = invalidTimeList
    end function

    '****************************************************************
    ' build a list of all available times for the selected day
    '****************************************************************
    private function GetTimeList(SelectedDate, NumberOfGuests)
        dim myDate : set myDate = new AvailableDate
        myDate.Load SelectedDate

        dim timeDictionary : set timeDictionary = GetFullTimeDictionary(myDate)
        dim eventsForTheDay : set eventsForTheDay = new ScheduledEventCollection
        eventsForTheDay.LoadByDate SelectedDate
    
        ' the following method updates the value in the time dictionary
        SetPlayerCount myDate, timeDictionary, eventsForTheDay

        dim softLimitList : set softLimitList = GetSoftLimitList(SelectedDate, timeDictionary, eventsForTheDay)

        set GetTimeList = GetFinalList(myDate, softLimitList, timeDictionary, NumberOfGuests)
    end function

    '****************************************************************
    ' count all of the players for every time slot in the day
    '****************************************************************
    private sub SetPlayerCount(EventDate, FullTimeDictionary, EventList)
        for each timeKey in FullTimeDictionary
            dim personCount: personCount = 0

            if bDebugMode then response.Write "DEBUG: time " & timeKey & "<br/>"

            for i = 0 to EventList.Count - 1
                if EventList(i).CanCountPlayers then
                    dim startDateTime : startDateTime = EventList(i).SelectedDate & " " & GetMilitaryTime(EventList(i).StartTime)
                    dim currentDateTime : currentDateTime = EventDate.SelectedDate & " " & timeKey
                    dim endDateTime : endDateTime = EventList(i).SelectedDate & " " & GetMilitaryTime(EventList(i).EstimatedEndTime)
                    dim startDiff : startDiff = DateDiff("n", startDateTime, currentDateTime)
                    dim endDiff : endDiff = DateDiff("n", endDateTime, currentDateTime)
         
                    if (startDiff >= 0 and endDiff <= 0 and len(EventList(i).NumberOfPatrons) > 0) then               
                        if bEditMode then
                            if EventList(i).EventId = myEvent.EventId then
                            else
                                personCount = personCount + cint(EventList(i).NumberOfPatrons)
                            end if
                        else
                            if bDebugMode then response.Write "DEBUG: " & GetMilitaryTime(EventList(i).StartTime) & " - " & EventList(i).PartyName & "(" & EventList(i).NumberOfPatrons & ")<br/>"
                            personCount = personCount + cint(EventList(i).NumberOfPatrons)
                        end if
                    end if
                end if
            next

            FullTimeDictionary(timeKey) = personCount
            if bDebugMode then response.Write "DEBUG: count = " & personCount & "<br/>"
        next
    end sub
%>