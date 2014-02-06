<!--#include file="../classes/includelist.asp"-->
<%
    Dim includeExpiredCookie, includeCheckedInCookie
    if IsEmpty(Request.Cookies("IncExpired")) OR IsNULL(Request.Cookies("IncExpired")) OR Len(Request.Cookies("IncExpired")) = 0 then 
        includeExpiredCookie = true 
    Else 
        includeExpiredCookie = Request.Cookies("IncExpired") 
    End If
    if IsEmpty(Request.Cookies("IncCheckedIn")) OR IsNULL(Request.Cookies("IncCheckedIn")) OR Len(Request.Cookies("IncCheckedIn")) = 0 then 
        includeCheckedInCookie = true 
    Else 
        includeCheckedInCookie = Request.Cookies("IncCheckedIn") 
    End If

    CONST ACTION_ADDNEWPLAYDATE = "59"
    CONST ACTION_ADMINCHECKINPLAYER = "60"
    CONST ACTION_ADMINCHECKOUTPLAYER = "61"
    CONST ACTION_BUILDSTREETFORDDL = "3"
    CONST ACTION_CHANGEGROUP = "58"
    CONST ACTION_CHANGEGROUPMULTIPLAYERS = "63"
    CONST ACTION_CHANGEPLAYDATE = "57"
    CONST ACTION_DELETEPLAYDATE = "62"
    const ACTION_DELETEWAIVER = "30"
    CONST ACTION_EMAILSEARCHWaiver = "2"
    CONST ACTION_GETPLAYERBYID = "5"
    CONST ACTION_GETWAIVERBYPID = "7"
    CONST ACTION_GETWAIVERLIST = "54"
    CONST ACTION_GETPLAYERSIG = "20"
    CONST ACTION_GETEVENTGROUPLIST = "53"
    const ACTION_GETGROUPDROPDOWNLIST = "4"
    const ACTION_GETWALKUPGROUP = "55"
    CONST ACTION_GETSEARCHRESULTS = "56"
    CONST ACTION_GETWAIVERLISTFORPRINT = "64"
    CONST ACTION_GETPLAYDATEBYPLAYERANDDATE = "65"
    const ACTION_SENDALLWAIVERS = "51"
    const ACTION_SENDALLWAIVERSBYPARENT = "52"
    const ACTION_SENDSINGLEWAIVER = "50"
    CONST ACTION_SEARCHFORWAIVERBYINFO = "1"
    CONST ACTION_SEARCHFORWAIVERBYNAMEADDY = "6"

    select case request.QueryString(QUERYSTRING_VAR_ACTION)
        case ACTION_ADMINCHECKINPLAYER
            AdminCheckInPlayer Request.QueryString(QUERYSTRING_VAR_PLAYERID), Request.QueryString(QUERYSTRING_VAR_PLAYDATE), Request.QueryString(QUERYSTRING_VAR_PLAYTIME)
        case ACTION_ADMINCHECKOUTPLAYER
            AdminCheckOUTPlayer Request.QueryString(QUERYSTRING_VAR_PLAYERID), Request.QueryString(QUERYSTRING_VAR_PLAYDATE), Request.QueryString(QUERYSTRING_VAR_PLAYTIME)
        case ACTION_ADDNEWPLAYDATE
            AddNewPlaydate Request.QueryString(QUERYSTRING_VAR_PLAYERID), Request.QueryString(QUERYSTRING_VAR_PLAYDATE), Request.QueryString(QUERYSTRING_VAR_PLAYTIME)
        case ACTION_BUILDSTREETFORDDL
            BuildStreetForDDL Request.QueryString(QUERYSTRING_VAR_PLAYERLIST)
        case ACTION_CHANGEGROUP
            ChangeGroup Request.QueryString(QUERYSTRING_VAR_PLAYERID), Request.QueryString(QUERYSTRING_VAR_PLAYDATE), Request.QueryString(QUERYSTRING_VAR_PLAYTIME), Request.QueryString(QUERYSTRING_VAR_EVENTID)
        case ACTION_CHANGEGROUPMULTIPLAYERS
            ChangeGroupMultiPlayers Request.QueryString(QUERYSTRING_VAR_PLAYERLIST), Request.QueryString(QUERYSTRING_VAR_PLAYDATE), Request.QueryString(QUERYSTRING_VAR_PLAYTIME), Request.QueryString(QUERYSTRING_VAR_EVENTID)
        case ACTION_CHANGEPLAYDATE
            ChangePlayDateTime Request.QueryString(QUERYSTRING_VAR_PLAYERID), Request.QueryString(QUERYSTRING_VAR_PLAYDATE), Request.QueryString(QUERYSTRING_VAR_PLAYTIME), Request.QueryString(QUERYSTRING_VAR_EVENTID), Request.QueryString(QUERYSTRING_VAR_ADMINCHECKIN), Request.QueryString(QUERYSTRING_VAR_OLDPLAYDATE), Request.QueryString(QUERYSTRING_VAR_OLDPLAYTIME)
        case ACTION_DELETEPLAYDATE
            DeletePlayDate Request.QueryString(QUERYSTRING_VAR_PLAYERID), Request.QueryString(QUERYSTRING_VAR_PLAYDATE), Request.QueryString(QUERYSTRING_VAR_PLAYTIME)
        case ACTION_DELETEWAIVER
            DeleteWaiver Request.QueryString(QUERYSTRING_VAR_WAIVERID)
        case ACTION_EMAILSEARCHWaiver
                SendWaiverSearchEmailResult Request.QueryString(QUERYSTRING_VAR_EMAILADDRESS)
        case ACTION_GETEVENTGROUPLIST
            BuildGroupList Request.QueryString(QUERYSTRING_VAR_SELECTEDDATE)
        case ACTION_GETGROUPDROPDOWNLIST
            BuildGroupDropDownList Request.QueryString(QUERYSTRING_VAR_SELECTEDDATE)
        case ACTION_GETPLAYDATEBYPLAYERANDDATE
            GetPlayDatesForSelectedDate Request.QueryString(QUERYSTRING_VAR_PLAYERID), Request.QueryString(QUERYSTRING_VAR_PLAYDATE)
        case ACTION_GETPLAYERBYID
            GetPlayersWaiverById Request.QueryString(QUERYSTRING_VAR_PLAYERID)
        case ACTION_GETPLAYERSIG
                WritePlayerSignature Request.QueryString(QUERYSTRING_VAR_WAIVERID)
        case ACTION_GETSEARCHRESULTS
            BuildSearchResults Request.QueryString(QUERYSTRING_VAR_TEXT)
        case ACTION_GETWAIVERBYPID
            GetWaiverIdByPlayerId Request.QueryString(QUERYSTRING_VAR_PLAYERID)
        case ACTION_GETWAIVERLIST
            BuildWaiverLIst Request.QueryString(QUERYSTRING_VAR_EVENTID), Request.QueryString(QUERYSTRING_VAR_SELECTEDDATE)
        case ACTION_GETWAIVERLISTFORPRINT
            BuildWaiverListForPrint Request.QueryString(QUERYSTRING_VAR_EVENTID), Request.QueryString(QUERYSTRING_VAR_SELECTEDDATE), Request.QueryString(QUERYSTRING_VAR_PLAYTIME)
        case ACTION_GETWALKUPGROUP
            BuildWalkupWaiverList Request.QueryString(QUERYSTRING_VAR_SELECTEDDATE), _
                                  Request.QueryString(QUERYSTRING_VAR_PLAYTIME)
        case ACTION_SENDALLWAIVERS
            SendAllWaiverEmail Request.QueryString(QUERYSTRING_VAR_WAIVERID)
        case ACTION_SENDALLWAIVERSBYPARENT
            SendAllWaiverEmailByParentEmailAddress Request.QueryString(QUERYSTRING_VAR_PLAYERID)
        case ACTION_SENDSINGLEWAIVER
            SendWaiverEmail Request.QueryString(QUERYSTRING_VAR_WAIVERID)
        case ACTION_SEARCHFORWAIVERBYINFO
                SearchForWaiverByInfo Request.QueryString(QUERYSTRING_VAR_FIRSTNAME), Request.QueryString(QUERYSTRING_VAR_LASTNAME), Request.QueryString(QUERYSTRING_VAR_DOB)
        case ACTION_SEARCHFORWAIVERBYNAMEADDY
                SearchForWaiverByNameDOBAddress Request.QueryString(QUERYSTRING_VAR_FIRSTNAME), Request.QueryString(QUERYSTRING_VAR_LASTNAME), Request.QueryString(QUERYSTRING_VAR_DOB), Request.QueryString(QUERYSTRING_VAR_ADRESS)

    end select

    '***********************************************************
    ' Add New Play Date 
    '***********************************************************
    private sub AddNewPlaydate(PlayerId, PlayDate, PlayTime)
        dim myPlayDates, matchFound, myNewPlayDate
        set myPlayDates = new PlayDateTimeCollection
        set myNewPlayDate = new PlayDateTime
        matchFound = false

        myPlayDates.LoadAllByPlayer PlayerId

        For ndx = 0 to myPlayDates.count -1
            If CDate(myPlayDates(ndx).PlayDate) = CDate(PlayDate) Then
                matchFound = true
            End If
        Next

        If matchFound Then
            response.write "Player already has a play date scheduled for the selected date."
        Else
            myNewPlayDate.PlayerId = PlayerId
            myNewPlayDate.PlayDate = PlayDate
            myNewPlayDate.PlayTime = PlayTime
            myNewPlayDate.EventId = NULL
            myNewPlayDate.CheckInDateTimeAdmin = NULL
       
            myNewPlayDate.AddNew
        End If
    end sub

    '***********************************************************
    ' Admin Check In Player
    '***********************************************************
    private sub AdminCheckInPlayer(PlayerId, PlayDate, PlayTime)
        dim myPlayDate, myNewPlayDate, myCheckInTime
        set myPlayDate = new PlayDateTimeCollection

        myPlayDate.LoadAllByPlayerAndDate PlayerId, Date
        myCheckInTime = GetTimeForAdminCheckIn()

        If myPlayDate.Count > 0 Then
            for index = 0 to myPlayDate.Count - 1
                If CDate(myPlayDate(index).PlayDate) = Date Then
                    myPlayDate(index).CheckInDateTimeAdmin = Now()
                    myPlayDate(index).UpdateAdminCheckIn   
                End If
            Next
        Else
            set myNewPlayDate = new PlayDateTime

            myNewPlayDate.PlayerId = PlayerId
            myNewPlayDate.PlayDate = Date
            myNewPlayDate.PlayTime = myCheckInTime
            myNewPlayDate.EventId = NULL
            myNewPlayDate.CheckInDateTimeAdmin = Now()

            myNewPlayDate.AddNew 
        End If

        response.write "true"
    end sub

    '***********************************************************
    ' Admin Check Out Player
    '***********************************************************
    private sub AdminCheckOutPlayer(PlayerId, PlayDate, PlayTime)
        dim myPlayDate
        set myPlayDate = new PlayDateTimeCollection

        myPlayDate.LoadAllByPlayerAndDate PlayerId, Date
        
        If myPlayDate.Count > 0 Then
            for index = 0 to myPlayDate.Count - 1      
                If CDate(myPlayDate(index).PlayDate) = Date AND NOT (IsNULL(myPlayDate(index).CheckInDateTimeAdmin) OR IsEmpty(myPlayDate(index).CheckInDateTimeAdmin) OR Len(myPlayDate(index).CheckInDateTimeAdmin) = 0) Then
                    myPlayDate(index).CheckInDateTimeAdmin = NULL
                    myPlayDate(index).UpdateAdminCheckIn 
                    response.write "Player has been checked out of today's play date."
                End If
            next
        Else
            response.write "Player does not have a play date for today."        
        End If    
    end sub

    '***********************************************************
    ' build a list that shows all of the events for the selecte date
    ' and check for waivers that aren't assigned to an event
    '***********************************************************
    private sub BuildGroupList(SelectedDate)
        dim myEvents, currentTime
        set myEvents = new ScheduledEventCollection

        Response.Write "<ul>"

        ' list all of the events for the day
        myEvents.LoadByDateWithWaivers SelectedDate

        currentTime = ""

        if myEvents.Count > 0 then
            for index = 0 to myEvents.Count - 1
                dim isWalkupGroup: isWalkupGroup = isnull(myEvents(index).EventId)

                if myEvents(index).StartTimeShortFormat <> currentTime then
                    currentTime = myEvents(index).StartTimeShortFormat
                    Response.Write "<li class=""timeentry"">" & currentTime & "</li>"
                end if

                if isWalkupGroup then
                    Response.Write "<li id=""myLine" & Replace(currentTime,":","") & """ onclick=""selectWalkupGroup('" & currentTime & "',this.id)"">"
                else
                    Response.Write "<li id=""myLine" & myEvents(index).EventId & """ onclick=""selectEvent(" & myEvents(index).EventId & ",this.id)"">"
                end if
            
                if len(myEvents(index).PartyName) > 0 then
                    response.Write myEvents(index).PartyName 
                else
                    response.Write myEvents(index).ContactName 
                end if

                if isWalkupGroup then
                    response.Write "(" & myEvents(index).WaiverCount & ")"
                else
                    ' build the count of waivers
                    'dim waivers
                    'set waivers = new WaiverCollection
                    'waivers.LoadByEventId myEvents(index).EventId
                
                    if len(myEvents(index).NumberOfPatrons) = 0 then myEvents(index).NumberOfPatrons = 0
                    'response.Write "(" & waivers.Count & "/" & myEvents(index).NumberOfPatrons & ")"
                    response.Write "(" & myEvents(index).WaiverCount & "/" & myEvents(index).NumberOfPatrons & ")"
                end if
                Response.Write"</li>"
            next

        end if

        ' list all of the ungrouped waivers for the day
        'dim ungroupedWaivers
        'set ungroupedWaivers = new WaiverCollection
        'ungroupedWaivers.LoadUngroupedByDate SelectedDate

        'if ungroupedWaivers.Count > 0 then
        '        Response.Write "<li onclick=""selectEvent('')"">"
        '        Response.Write "Walk-up Players(" & ungroupedWaivers.Count & ")"
        '        Response.Write"</li>"
        'end if

        Response.Write "</ul>"
    end sub
    
    '************************************************************
    ' build a drop-down list of all events for the day
    '************************************************************
    private sub BuildGroupDropDownList(SelectedDate)
        dim myEvents
        set myEvents = new ScheduledEventCollection

        response.Write "<option value="""" selected=""selected"">Please Select</option>" & vbCrLf
        
        ' list all of the events for the day
        myEvents.LoadValidByDate SelectedDate

        for index = 0 to myEvents.Count - 1
            response.Write "<option value=""" & myEvents(index).EventId  & """>" & GetTimeString(myEvents(index).StartTimeShortFormat) & " "

            if len(myEvents(index).PartyName) > 0 then
                response.Write myEvents(index).PartyName 
            else
                response.Write myEvents(index).ContactName 
            end if

            response.Write "</option>" & vbCrLf
        next

        response.Write "<option value=""-1"">Group not listed</option>" & vbCrLf
    end sub

    '***********************************************************
    ' build a list of all the search results
    '***********************************************************
    private sub BuildSearchResults(SearchText)
        dim waivers, myPlayDate
        set waivers = new WaiverCollection

        waivers.Search SearchText, includeExpiredCookie, includeCheckedInCookie

            Response.write "<ul  class=""menuBarWaiverList"" style=""display: block"">"
            Response.write "<li>"
            Response.write "Include Expired <input id=""cbxIncExpired"" name=""cbxIncExpired"" type=""checkbox"" onClick=""setExpiredFilter();"" "
            If includeExpiredCookie Then Response.Write "checked" End If
            response.write " />"
            Response.write "</li>"
            Response.write "<li></li>"
            Response.write "<li>"
            Response.write "Include Checked In <input id=""cbxIncCheckedIn"" name=""cbxIncCheckedIn"" type=""checkbox"" onClick=""setCheckedInFilter();"" "
            If includeCheckedInCookie Then Response.Write "checked" End If
            response.write " />"
            Response.write "</li>"
            Response.write "</ul>"

StartTimer 1

        if waivers.Count > 0 then

                Response.write "<ul class=""menuBarWaiverList"" style=""display: block"">"
                Response.write "<li onClick=""checkInMulti('search')"">"
                Response.write "<img src=""../../content/images/door_in.png"" alt="""" title=""Check In Players"" border=""0"" />"
                Response.write " Check In"
                Response.write "</li>"
                Response.write "<li onClick=""checkOutMulti('search')"">"
                Response.write "<img src=""../../content/images/door_out.png"" alt="""" title=""Check Out Players"" border=""0"" />"
                Response.write " Check Out"
                Response.write "</li>"
                Response.write "<li onClick=""showEventGroupsForMultiselect()"">"
                Response.write "<img src=""../../content/images/group.png"" alt="""" title=""Change Players Group"" border=""0"" />"
                Response.write " Change Group"
                Response.write "</li>"
                Response.write "</ul>" 

            Response.Write "<table class=""results"">"
            response.Write "<thead>"
            response.Write "<tr>"
            response.Write "<th/>"
            response.Write "<th style=""text-align: left;"">Name</th>"
            response.Write "<th style=""text-align: left;"">Age</th>"
            response.Write "<th colspan=2 style=""text-align: left;"">DOB</th>"
            response.Write "</tr>"
            response.Write "</thead>"
            response.Write "<tbody>"
            
            for i = 0 to waivers.Count - 1

                If waivers(i).ValidityStatus = WAIVER_VALIDITYSTATUS_GOOD Then 
                    Response.Write "<tr class=""resultsDetail"" onclick=""selectWaiver('" & waivers(i).HashId & "', '" & waivers(i).CreateDateTime &"')"">"
                Else
                    Response.Write "<tr style=""color:red;"" class=""resultsDetail"" onclick=""selectWaiver('" & waivers(i).HashId & "', '" & waivers(i).CreateDateTime &"')"">"
                End If         

                response.Write "<td style=""text-align: left;"">"

                If waivers(i).CheckedInToday AND waivers(i).ValidityStatus = WAIVER_VALIDITYSTATUS_GOOD Then
                   response.write "<img src=""../../content/images/accept.png"" alt=""Checked In"" title=""Checked In"" /></td>"
                Else
                   response.write "&nbsp;</td>"
                End If

                response.Write "<td>" & waivers(i).LastName & ", " & waivers(i).FirstName & "</td>"
                response.Write "<td class=""center"">" & waivers(i).CurrentAge & "</td>"
                response.Write "<td class=""right"">" & FormatDateTime(waivers(i).PlayerDOB, 2) & "</td>"
                response.write "<td>"
                If waivers(i).IsValid Then
                    response.Write "<input type=""checkbox"" id=""txtCheckMeIn" & i & """ name=""txtCheckMeIn" & i & """ value=""pd=" & waivers(i).PlayerId & "&pdt=" & FormatDateTime(Date,2) & "&pt=" & FormatDateTime(Time,4) & """"
                End If
                response.write "</td>"
                Response.Write"</tr>"
            next

            response.Write "</tbody>"
            Response.Write "</table>"
            response.write "<table style=""width: 100%""><tr><td colspan=2>&nbsp;</td></tr><td colspan=2>&nbsp;</td></tr>"
            response.Write "<tr><td>&nbsp;</td>"
            response.write "    <td class=""right""><input style=""width:100px"" type=""button"" name=""btnSelectAll"" value=""Select All"" onClick=""selectAllPlayers(this.value);"" /></td></tr>"
            response.write "<tr><td colspan=2 class=""right""><input style=""width:100px"" type=""button"" name=""btnSelectAll"" value=""Unselect All"" onClick=""selectAllPlayers(this.value);"" /></td></tr>"
            response.write "</table>"
        end if
            Response.write "<input id=""txtSelectedPlayTime"" type=""hidden"" value=""" & Time & """ />"


'        ElapsedLap(1) = StopTimer(1)
'        StartTimer 2

ElapsedLap(1) = StopTimer(1)
TotalElapsed = TotalElapsed + ElapsedLap(1)

response.write "<br />Total Records: " & waivers.count
response.write "<br />Time Total: " & TotalElapsed    
response.write "<br />"
'response.write "<br />Time 1: " & ElapsedLap(0)    
response.write "<br />Time 2: " & ElapsedLap(1)    


    end sub

    '***********************************************************
    ' build drop-down list for street chooser
    '************************************************************
    private sub BuildStreetForDDL(listOfPlayers)
        dim streets, myList, myPlayer, newList(2,10), randomNdx

        set streets = new PlayerCollection
        streets.LoadPlayersStreetChooser()
 
        myList = Split(listOfPlayers,"@")

        for y = 0 to streets.count -1
            newList(0,y) = streets(y).playerid
            newList(1,y) = Right(streets(y).address,Len(streets(y).address) - InStr(streets(y).address," "))
            newList(2,y) = "no"
        next        
        set streets = nothing
 
        for each x in myList
            set myPlayer = new Player
            if x <> "" then
                myPlayer.LoadById x
                do 
                    Randomize
                    randomNdx = Int(Rnd * UBound(newList,2)) 
                loop until newList(2,randomNdx) = "no"   

                newList(0,randomNdx) = myPlayer.playerId 
                newList(1,randomNdx) = Right(myPlayer.address,Len(myPlayer.address) - InStr(myPlayer.address," ")) 
                newList(2,randomNdx) = "yes"

               set myPlayer = nothing
            end if
        next    

        for i = 0 to Ubound(newList,2) -1
            response.write "<input type=""radio"" name=""street-radio"" id=""street-radio-" & i & """ value=""" & newList(0,i) & """ />"
            response.write "<label for=""street-radio-" & i & """>" & newList(1,i) & "</label>"
        next     
    
        response.write "<input type=""radio"" name=""street-radio"" id=""street-radio-" & ubound(newList,2) & """ value=""street not listed"" />"
        response.write "<label for=""street-radio-" & i & """>Street not listed</label>"
              

    end sub

    '***********************************************************
    ' build a list of all waivers for the event or date
    '***********************************************************
    private sub BuildWaiverList(EventId, SelectedDate)
        dim waivers, myPlayDate, myEvent
        set waivers = new WaiverCollection
        set myEvent = new ScheduledEvent

        if len(EventId) = 0 then
            waivers.LoadUngroupedByDate SelectedDate, includeExpiredCookie, includeCheckedInCookie
        else
            waivers.LoadByEventId EventId, includeExpiredCookie, includeCheckedInCookie
            myEvent.Load EventId
            PlayTime = myEvent.StartTime
        end if

            Response.write "<ul  class=""menuBarWaiverList"" style=""display: block"">"
            Response.write "<li>"
            Response.write "Include Expired <input id=""cbxIncExpired"" name=""cbxIncExpired"" type=""checkbox"" onClick=""setExpiredFilter();"" "
            If includeExpiredCookie Then Response.Write "checked" End If
            response.write " />"
            Response.write "</li>"
            Response.write "<li></li>"
            Response.write "<li>"
            Response.write "Include Checked In <input id=""cbxIncCheckedIn"" name=""cbxIncCheckedIn"" type=""checkbox"" onClick=""setCheckedInFilter();"" "
            If includeCheckedInCookie Then Response.Write "checked" End If
            response.write " />"
            Response.write "</li>"
            Response.write "</ul>"

        if waivers.Count > 0 then

             If CDate(SelectedDate) = Date Then
                Response.write "<ul class=""menuBarWaiverList"" style=""display: block"">"
                Response.write "<li onClick=""checkInMulti('waiverlist')"">"
                Response.write "<img src=""../../content/images/door_in.png"" alt="""" title=""Check In Players"" border=""0"" />"
                Response.write " Check In"
                Response.write "</li>"
                Response.write "<li onClick=""checkOutMulti('waiverlist')"">"
                Response.write "<img src=""../../content/images/door_out.png"" alt="""" title=""Check Out Players"" border=""0"" />"
                Response.write " Check Out"
                Response.write "</li>"
                Response.write "<li onClick=""showEventGroupsForMultiselect()"">"
                Response.write "<img src=""../../content/images/group.png"" alt="""" title=""Change Players Group"" border=""0"" />"
                Response.write " Change Group"
                Response.write "</li>"
                Response.write "</ul>"    
            End If        

            Response.Write "<table class=""results"">"
            response.Write "<thead>"
            response.Write "<tr>"
            response.Write "<th/>"
            response.Write "<th style=""text-align: left;"">Name</th>"
            response.Write "<th style=""text-align: left;"">Age</th>"
            response.Write "<th colspan=2 style=""text-align: left;"">DOB</th>"
            response.Write "</tr>"
            response.Write "</thead>"
            response.Write "<tbody>"

            for i = 0 to waivers.Count - 1
                set myPlayDate = new PlayDateTime

                myPlayDate.Load waivers(i).PlayerId, SelectedDate, PlayTime

                If waivers(i).ValidityStatus = WAIVER_VALIDITYSTATUS_GOOD Then 
                    Response.Write "<tr class=""resultsDetail"" onclick=""selectWaiver('" & waivers(i).HashId & "', '" & SelectedDate &"')"">"
                Else
                    Response.Write "<tr style=""color:red;"" class=""resultsDetail"" onclick=""selectWaiver('" & waivers(i).HashId & "', '" & SelectedDate &"')"">"
                End If         
            
                response.Write "<td style=""text-align: left;"">"

                If IsEmpty(myPlayDate.CheckInDateTimeAdmin) OR _
                   IsNull(myPlayDate.CheckInDateTimeAdmin) OR _
                   Len(myPlayDate.CheckInDateTimeAdmin) <= 0 _
                Then
                   response.write "&nbsp;</td>"
                Else
                   response.write "<img src=""../../content/images/accept.png"" alt=""Checked In"" title=""Checked In"" /></td>"
                End If

                response.write "<td>" & waivers(i).LastName & ", " & waivers(i).FirstName & "</td>"
                response.Write "<td class=""center"">" & waivers(i).CurrentAge & "</td>"
                response.Write "<td class=""right"">" & FormatDateTime(waivers(i).PlayerDOB, 2) & "</td>"
                response.write "<td>"
                If waivers(i).IsValid Then
                    response.Write "<input type=""checkbox"" id=""txtCheckMeIn" & i & """ name=""txtCheckMeIn" & i & """ value=""pd=" & waivers(i).PlayerId & "&pdt=" & SelectedDate & "&pt=" & Left(PlayTime,5) & """"
                End If
                response.write "</td>"
                response.write "</tr>"      

                set myPlayDate = Nothing
            next
            
            response.Write "</tbody>"
            Response.Write "</table>"
            response.write "<table style=""width: 100%""><tr><td colspan=2>&nbsp;</td></tr><td colspan=2>&nbsp;</td></tr>"
            response.Write "<tr><td><a href=""printWaiverList.asp?ev=" & EventId & "&dt=" & SelectedDate & "&pt=" & PlayTime  & """ target=""_blank"">Printable Version</a></td>"
            response.write "    <td class=""right""><input style=""width:100px"" type=""button"" name=""btnSelectAll"" value=""Select All"" onClick=""selectAllPlayers(this.value);"" /></td></tr>"
            response.write "<tr><td colspan=2 class=""right""><input style=""width:100px"" type=""button"" name=""btnSelectAll"" value=""Unselect All"" onClick=""selectAllPlayers(this.value);"" /></td></tr>"
            response.write "</table>"
        end if
            Response.write "<input id=""txtSelectedPlayTime"" name=""txtSelectedPlayTime"" type=""hidden"" value=""" &  Left(PlayTime,5) & """ />"
    end sub
    
    '***********************************************************
    ' build a list of all waivers for the event or date
    '***********************************************************
    private sub BuildWaiverListForPrint(EventId, SelectedDate, PlayTime)
        dim waivers, myPlayDate, myEvent
        set waivers = new WaiverCollection
        set myEvent = new ScheduledEvent

        if len(EventId) = 0 then
            waivers.LoadUngroupedByDate SelectedDate, includeExpiredCookie, includeCheckedInCookie
        else
            waivers.LoadByEventId EventId, includeExpiredCookie, includeCheckedInCookie
            myEvent.Load EventId
            PlayTime = myEvent.StartTime
        end if

        if waivers.Count > 0 then
            Response.write SelectedDate & "&nbsp;&nbsp;" & Left(PlayTime,5) & "&nbsp;&nbsp;"
            If len(EventId) <> 0 Then Response.Write myEvent.PartyName & "<br />" Else Response.Write "Walk-Up Players<br />" End If
            Response.Write "<table class=""results"">"
            response.Write "<thead>"
            response.Write "<tr>"
            response.Write "<th/>"
            response.Write "<th style=""text-align: left;"">Name</th>"
            response.Write "<th style=""text-align: left;"">Age</th>"
            response.Write "<th colspan=2 style=""text-align: left;"">DOB</th>"
            response.Write "</tr>"
            response.Write "</thead>"
            response.Write "<tbody>"

            for i = 0 to waivers.Count - 1
                set myPlayDate = new PlayDateTime

                myPlayDate.Load waivers(i).PlayerId, SelectedDate, PlayTime

                If waivers(i).ValidityStatus = WAIVER_VALIDITYSTATUS_GOOD Then 
                    Response.Write "<tr class=""resultsDetail"" onclick=""selectWaiver('" & waivers(i).HashId & "', '" & SelectedDate &"')"">"
                Else
                    Response.Write "<tr style=""color:red;"" class=""resultsDetail"" onclick=""selectWaiver('" & waivers(i).HashId & "', '" & SelectedDate &"')"">"
                End If         
            
                response.Write "<td style=""text-align: left;"">"

                If IsEmpty(myPlayDate.CheckInDateTimeAdmin) OR _
                   IsNull(myPlayDate.CheckInDateTimeAdmin) OR _
                   Len(myPlayDate.CheckInDateTimeAdmin) <= 0 _
                Then
                   response.write "&nbsp;</td>"
                Else
                   response.write "<img src=""../../content/images/accept.png"" alt=""Checked In"" title=""Checked In"" /></td>"
                End If

                response.write "<td>" & waivers(i).LastName & ", " & waivers(i).FirstName & "</td>"
                response.Write "<td class=""center"">" & waivers(i).CurrentAge & "</td>"
                response.Write "<td class=""right"">" & FormatDateTime(waivers(i).PlayerDOB, 2) & "</td>"
                response.write "<td/>"
                response.write "</tr>"      

                set myPlayDate = Nothing
            next
            
            response.Write "</tbody>"
            Response.Write "</table>"
            response.write "<table style=""width: 100%""><tr><td>&nbsp;</td></tr><td>&nbsp;</td></tr>"
'            response.Write "<tr><td><a href=""printWaiverList.asp?ev=" & EventId & "&dt=" & SelectedDate & """ target=""_blank"">Printable Version</a></td></tr>"
            response.write "</table>"

        end if
    end sub

    '***********************************************************
    ' build a list of all waivers for the event or date
    '***********************************************************
    private sub BuildWalkupWaiverList(SelectedDate, PlayTime)
        dim waivers, myPlayDate, IncludeInvalid
        set waivers = new WaiverCollection

        waivers.LoadUngroupedByDateAndTime SelectedDate, PlayTime, includeExpiredCookie, includeCheckedInCookie

            Response.write "<ul  class=""menuBarWaiverList"" style=""display: block"">"
            Response.write "<li>"
            Response.write "Include Expired <input id=""cbxIncExpired"" name=""cbxIncExpired"" type=""checkbox"" onClick=""setExpiredFilter();"" "
            If includeExpiredCookie Then Response.Write "checked" End If
            response.write " />"
            Response.write "</li>"
            Response.write "<li></li>"
            Response.write "<li>"
            Response.write "Include Checked In <input id=""cbxIncCheckedIn"" name=""cbxIncCheckedIn"" type=""checkbox"" onClick=""setCheckedInFilter();"" "
            If includeCheckedInCookie Then Response.Write "checked" End If
            response.write " />"
            Response.write "</li>"
            Response.write "</ul>"

        if waivers.Count > 0 then
            If CDate(SelectedDate) = Date Then
                Response.write "<ul class=""menuBarWaiverList"" style=""display: block"">"
                Response.write "<li onClick=""checkInMulti('waiverlist')"">"
                Response.write "<img src=""../../content/images/door_in.png"" alt="""" title=""Check In Players"" border=""0"" />"
                Response.write " Check In"
                Response.write "</li>"
                Response.write "<li onClick=""checkOutMulti('waiverlist')"">"
                Response.write "<img src=""../../content/images/door_out.png"" alt="""" title=""Check Out Players"" border=""0"" />"
                Response.write " Check Out"
                Response.write "</li>"
                Response.write "<li onClick=""showEventGroupsForMultiselect()"">"
                Response.write "<img src=""../../content/images/group.png"" alt="""" title=""Change Players Group"" border=""0"" />"
                Response.write " Change Group"
                Response.write "</li>"
                Response.write "</ul>" 
            End If
        
            Response.Write "<table class=""results"">"
            response.Write "<thead>"
            response.Write "<tr>"
            response.Write "<th/>"
            response.Write "<th style=""text-align: left;"">Name</th>"
            response.Write "<th style=""text-align: left;"">Age</th>"
            response.Write "<th colspan=2 style=""text-align: left;"">DOB</th>"
            response.Write "</tr>"
            response.Write "</thead>"
            response.Write "<tbody>"
            
            for i = 0 to waivers.Count - 1
                set myPlayDate = new PlayDateTime
                myPlayDate.Load waivers(i).PlayerId, SelectedDate, PlayTime
              
                If waivers(i).ValidityStatus = WAIVER_VALIDITYSTATUS_GOOD Then 
                    Response.Write "<tr class=""resultsDetail"" onclick=""selectWaiver('" & waivers(i).HashId & "', '" & SelectedDate &"')"">"
                Else
                    Response.Write "<tr  style=""color:red;"" class=""resultsDetail"" onclick=""selectWaiver('" & waivers(i).HashId & "', '" & SelectedDate &"')"">"
                End If         

                response.Write "<td style=""text-align: left;"">"

                If IsEmpty(myPlayDate.CheckInDateTimeAdmin) OR _
                   IsNull(myPlayDate.CheckInDateTimeAdmin) OR _
                   Len(myPlayDate.CheckInDateTimeAdmin) <= 0 _
                Then
                   response.write "</td>"
                Else
                   response.write "<img src=""../../content/images/accept.png"" alt=""Checked In"" title=""Checked In"" /></td>"
                End If
                
                response.Write "<td>" & waivers(i).LastName & ", " & waivers(i).FirstName & "</td>"
                response.Write "<td class=""center"">" & waivers(i).CurrentAge & "</td>"
                response.Write "<td class=""right"">" & FormatDateTime(waivers(i).PlayerDOB, 2) & "</td>"
        
                response.write "<td>"
                If waivers(i).IsValid Then
                    response.Write "<input type=""checkbox"" id=""txtCheckMeIn" & i & """ name=""txtCheckMeIn" & i & """ value=""pd=" & waivers(i).PlayerId & "&pdt=" & SelectedDate & "&pt=" & Left(myPlayDate.PlayTime,5) & """"
                End If
                response.write "</td>"
                response.write "</tr>"    

                set myPlayDate = Nothing
            next

            response.Write "</tbody>"
            Response.Write "</table>"
            response.write "<table style=""width: 100%""><tr><td colspan=2>&nbsp;</td></tr><td colspan=2>&nbsp;</td></tr>"
            response.Write "<tr><td><a href=""printWaiverList.asp?ev=" & EventId & "&dt=" & SelectedDate & "&pt=" & PlayTime  & """ target=""_blank"">Printable Version</a></td>"
            response.write "    <td class=""right""><input style=""width:100px"" type=""button"" name=""btnSelectAll"" value=""Select All"" onClick=""selectAllPlayers(this.value);"" /></td></tr>"
            response.write "<tr><td colspan=2 class=""right""><input style=""width:100px"" type=""button"" name=""btnSelectAll"" value=""Unselect All"" onClick=""selectAllPlayers(this.value);"" /></td></tr>"
            response.write "</table>"
        end if
            Response.write "<input id=""txtSelectedPlayTime"" type=""hidden"" value=""" &  Left(PlayTime,5) & """ />"
    end sub

    '***********************************************************
    ' Change Play Date
    '***********************************************************
    private sub ChangePlayDateTime(PlayerId, newPlayDate, newPlayTime, EventId, CheckInAdmin, originalPlayDate, originalPlayTime)
        dim myPlayDates, matchFound, myoriginalPlayDate
        set myPlayDates = new PlayDateTimeCollection
        set myoriginalPlayDate = new PlayDateTime

        matchFound = false

        If EventId = "" or IsNULL(EventId) Then EventId = NULL End If
        If originalPlayDate = "" or IsNULL(originalPlayDate) Then originalPlayDate = newPlayDate    
        If originalPlayTime = "" or IsNULL(originalPlayTime) Then originalPlayTime = newPlayTime    
        If CheckInAdmin = "" or IsNULL(CheckInAdmin) Then CheckInAdmin = NULL    

        newPlayTime = Left(newPlayTime,8)
        originalPlayTime = Left(originalPlayTime,8)

        myoriginalPlayDate.Load PlayerId, originalPlayDate, originalPlayTime  

        myoriginalPlayDate.PlayTime = newPlayTime          
        myoriginalPlayDate.EventId = EventId
        myoriginalPlayDate.CheckInDateTimeAdmin = CheckInAdmin

        'Changing to a different day
            If CDate(originalPlayDate) <> CDate(newPlayDate) Then
                If CDate(newPlayDate) < Date Then
                    response.write "Cannot edit the past."
                Else
                    myPlayDates.LoadAllByPlayer PlayerId
                    For ndx = 0 to myPlayDates.count -1
                        If CDate(myPlayDates(ndx).PlayDate) = CDate(newPlayDate) Then
                            matchFound = true
                        End If
                    Next

                    If matchFound Then
                        response.write "Player already has a play date scheduled for the selected date."
                    Else
                         myoriginalPlayDate.EventId = NULL 
                         myoriginalPlayDate.PlayDate = newPlayDate
                         myoriginalPlayDate.UpdatePlayDateTime originalPlayDate, originalPlayTime
                    End If
                End If
            Else
                myoriginalPlayDate.PlayDate = newPlayDate
                myoriginalPlayDate.UpdatePlayDateTime originalPlayDate, originalPlayTime
            End If
    end sub
    
    '***********************************************************
    ' Change Group Single Player
    '***********************************************************
    private sub ChangeGroup(PlayerId, PlayDate, PlayTime, EventId)
        dim myPlayDates, myEvent, oldPlayTime, newPlayTime
        set myPlayDates = new PlayDateTimeCollection
        set myEvent = new ScheduledEvent

        myPlayDates.LoadAllByPlayerAndDate PlayerId, Date

        If myPlayDates.Count > 0 Then
            for index = 0 to myPlayDates.Count - 1
                If CDate(myPlayDates(index).PlayDate) = Date Then
                    if len(EventId) > 0 then
                        myEvent.Load EventId
                        newPlayTime = myEvent.StartTimeShortFormat
                        oldPlayTime = myPlayDates(index).PlayTime
                    Else
                        EventId = NULL
                        newPlayTime = GetTimeForAdminCheckIn()
                        oldPlayTime = myPlayDates(index).PlayTime
                    End If

                    myPlayDates(index).EventId = EventId
                    myPlayDates(index).PlayTime = newPlayTime
                    myPlayDates(index).UpdatePlayDateEventAndTime oldPlayTime

                End If
            Next
        Else
            set myNewPlayDate = new PlayDateTime

            myNewPlayDate.PlayerId = PlayerId
            myNewPlayDate.PlayDate = Date
            myNewPlayDate.PlayTime = myCheckInTime
            myNewPlayDate.EventId = NULL
            myNewPlayDate.CheckInDateTimeAdmin = Now()

            myNewPlayDate.AddNew 
        End If
    end sub

    '***********************************************************
    ' Change Group For Multiple Players
    '***********************************************************
    private sub ChangeGroupMultiPlayers(PlayerList, PlayDate, PlayTime, EventId)
        dim myPlayDates, myNewPlayDate, myEvent, oldPlayTime, myPlayers, newPlayTime
        set myEvent = new ScheduledEvent

        myPlayers = split(PlayerList,"@")
        for each pId in myPlayers
            If pId <> "" Then
                set myPlayDates = new PlayDateTimeCollection
                set myEvent = new ScheduledEvent

                myPlayDates.LoadAllByPlayerAndDate pId, Date

                If myPlayDates.Count > 0 Then
                    for index = 0 to myPlayDates.Count - 1
                        If CDate(myPlayDates(index).PlayDate) = Date Then
                            if len(EventId) > 0 then
                                myEvent.Load EventId
                                newPlayTime = myEvent.StartTimeShortFormat
                                oldPlayTime = myPlayDates(index).PlayTime
                            Else
                                EventId = NULL
                                newPlayTime = GetTimeForAdminCheckIn()
                                oldPlayTime = myPlayDates(index).PlayTime
                            End If

                            myPlayDates(index).EventId = EventId
                            myPlayDates(index).PlayTime = newPlayTime
                            myPlayDates(index).UpdatePlayDateEventAndTime oldPlayTime

                        End If
                    Next
                Else
                    if len(EventId) > 0 then
                        myEvent.Load EventId
                        newPlayTime = myEvent.StartTimeShortFormat
                    Else
                        EventId = NULL
                        newPlayTime = GetTimeForAdminCheckIn()
                    End If

                    set myNewPlayDate = new PlayDateTime

                    myNewPlayDate.PlayerId = pId
                    myNewPlayDate.PlayDate = Date
                    myNewPlayDate.PlayTime = newPlayTime
                    myNewPlayDate.EventId = EventId
                    myNewPlayDate.CheckInDateTimeAdmin = Now()

                    myNewPlayDate.AddNew 
                End If                
            End If
        Next
    end sub

    '***********************************************************
    ' Delete Play Date
    '***********************************************************
    private sub DeletePlayDate(PlayerId, PlayDate, PlayTime)
        dim myPlayDate
        set myPlayDate = new PlayDateTime
        
        myPlayDate.PlayerId = PlayerId
        myPlayDate.PlayDate = PlayDate
        myPlayDate.PlayTime = PlayTime

        myPlayDate.Delete  

    end sub

    '***********************************************************
    ' Delete Waiver
    '***********************************************************
    private sub DeleteWaiver(WaiverHashId)
        dim myWaiver
        set myWaiver = new Waiver

        myWaiver.Delete WaiverHashId
    end sub
    
    '**********************************************************
    ' Get player's playdates for selected date 
    '**********************************************************
    sub GetPlayDatesForSelectedDate(PlayerId, SelectedDate)
        dim myPlayDates
        set myPlayDates = new PlayDateTimeCollection
 
        myPlayDates.LoadAllByPlayerAndDate PlayerId, SelectedDate

        if myPlayDates.Count > 0 Then
            response.write "[PLAYDATESFOUND]"
        else
            response.write "[NOPLAYDATES]"
        end If
    end sub

    '***********************************************************
    ' get players waivers status by id
    '************************************************************
    private sub GetPlayersWaiverById(nPlayerId)
        dim player, waivers, foundGoodWaiver, foundInvalidWaiver
        set player = new Player

        player.LoadById nPlayerId

        set waivers = new WaiverCollection
        waivers.LoadWaiversByPlayerId player.playerId

        if waivers.count > 0 Then
            foundGoodWaiver = false
            foundInvalidWaiver = false

            for ndx = 0 to waivers.count -1
                if waivers(ndx).ValidityStatus = WAIVER_VALIDITYSTATUS_GOOD Then
                    foundGoodWaiver = true
                end If
                if waivers(ndx).ValidityStatus = WAIVER_VALIDITYSTATUS_EXPIRED Then
                    foundInvalidWaiver = true
                end If
                if waivers(ndx).ValidityStatus = WAIVER_VALIDITYSTATUS_NEWADULT Then
                    foundInvalidWaiver = true
                end If
           next
      end if
                
      if foundGoodWaiver then
           response.write "1" & nPlayerId
      else
           if foundInvalidWaiver then
              response.write "3" & nPlayerId
           else
              response.write "4" & nPlayerId
           end if
      end if
    
    end sub


    '***********************************************************
    ' Get waiver id by player id
    '************************************************************
    
    private sub GetWaiverIdByPlayerId(playerid)
        dim myWaivers, foundWaiver
        set myWaivers = new WaiverCollection
        foundWaiver = false

        myWaivers.LoadWaiversByPlayerId playerId

        for i = 0 to myWaivers.Count -1
            if myWaivers(i).ValidityStatus = WAIVER_VALIDITYSTATUS_GOOD then
                response.write myWaivers(i).WaiverId
                foundWaiver = true
            end if
        next

        if NOT foundWaiver then
            response.write "waiver not found"
        end if
    end sub

    '***********************************************************
    ' Search for waivers by info
    '************************************************************
    private sub SearchForWaiverByInfo(sFirstName, sLastName, dtDOB)
        dim players, waivers, foundGoodWaiver, foundInvalidWaiver
        set players = new PlayerCollection

        players.LoadPlayersByNameDOB sFirstName, sLastName, dtDOB

        Select Case True
            Case players.Count = 1
                    set waivers = new WaiverCollection
                    waivers.LoadWaiversByPlayerId players(0).playerId

                    if waivers.count > 0 Then
                        foundGoodWaiver = false
                        foundInvalidWaiver = false

                        for ndx = 0 to waivers.count -1
                            if waivers(ndx).ValidityStatus = WAIVER_VALIDITYSTATUS_GOOD Then
                                foundGoodWaiver = true
                            end If
                            if waivers(ndx).ValidityStatus = WAIVER_VALIDITYSTATUS_EXPIRED Then
                                foundInvalidWaiver = true
                            end If
                            if waivers(ndx).ValidityStatus = WAIVER_VALIDITYSTATUS_NEWADULT Then
                                foundInvalidWaiver = true
                            end If
                        next
                    end if
                
                    if foundGoodWaiver then
                        response.write "1" & players(0).playerId
                    else
                        if foundInvalidWaiver then
                            response.write "3" & players(0).playerId
                        else
                            response.write "4" & players(0).playerId
                        end if
                    end if
   
            Case players.Count >= 2
                    response.write "2"
                    for x = 0 to players.count -1
                        response.write "@" & players(x).playerId
                    next
            Case Else
                    response.write "no players found"
        End Select
    end sub


    '***********************************************************
    ' Search for waivers by name dob address
    '************************************************************
    private sub SearchForWaiverByNameDOBAddress(sFirstName, sLastName, dtDOB, sAddress)
        dim players, waivers, foundGoodWaiver, foundInvalidWaiver
        set players = new PlayerCollection

        players.LoadPlayersByNameDOBAddress sFirstName, sLastName, dtDOB, sAddress

        Select Case True
            Case players.Count = 1
                    set waivers = new WaiverCollection
                    waivers.LoadWaiversByPlayerId players(0).playerId

                    if waivers.count > 0 Then
                        foundGoodWaiver = false
                        foundInvalidWaiver = false

                        for ndx = 0 to waivers.count -1
                            if waivers(ndx).ValidityStatus = WAIVER_VALIDITYSTATUS_GOOD Then
                                foundGoodWaiver = true
                            end If
                            if waivers(ndx).ValidityStatus = WAIVER_VALIDITYSTATUS_EXPIRED Then
                                foundInvalidWaiver = true
                            end If
                            if waivers(ndx).ValidityStatus = WAIVER_VALIDITYSTATUS_NEWADULT Then
                                foundInvalidWaiver = true
                            end If
                        next
                    end if
                
                    if foundGoodWaiver then
                        response.write "1" & players(0).playerId
                    else
                        if foundInvalidWaiver then
                            response.write "3" & players(0).playerId
                        else
                            response.write "4" & players(0).playerId
                        end if
                    end if
   
            Case players.Count >= 2
                    response.write "2"
                    for x = 0 to players.count -1
                        response.write "@" & players(x).playerId
                    next
            Case Else
                    response.write "no players found"
        End Select
    end sub


    '***********************************************************
    ' Send waiver search email
    '************************************************************
    private sub SendWaiverSearchEmailResult(email)
        dim waivers, trailingS
        set waivers = new WaiverCollection

        waivers.LoadWaiversByEmail email
        
        if waivers.Count > 0 then
            'Waiver(s) found - Send email

            dim em
            set em = new WaiverSearchEmail

            em.SendWaiverList waivers, email

            'Waivers(s) found - let users know
            if waivers.count > 1 then
                trailingS = "s"
            else
                trailingS = ""
            end if
                response.write "<br />We found " & waivers.Count & " waiver" & trailingS & " that match the provided email address.<br /><br />Please check your email to view the waiver" & trailingS & ".<br /><br />" 
        else
            response.write "no email found for waivers"
        end if
    end sub
        
    private sub SendAllWaiverEmailByParentEmailAddress(playerID)
        dim parentPlayer, waivers
        set parentPlayer = new Player
        set waivers = new WaiverCollection

        parentPlayer.LoadById playerID
        waivers.LoadWaiversByEmail parentplayer.emailaddress

        dim em
        set em = new WaiverSearchEmail

        em.SendWaiverList waivers, parentPlayer.emailaddress
    end sub

    private sub SendAllWaiverEmail(waiverhashid)
        dim parentWaiver, waivers
        set parentWaiver = new Waiver
        set waivers = new WaiverCollection

        parentWaiver.Load waiverhashid
        waivers.LoadWaiversByEmail parentWaiver.emailaddress

        dim em
        set em = new WaiverSearchEmail

        em.SendWaiverList waivers, parentWaiver.emailaddress
    end sub

    private sub SendWaiverEmail(WaiverHashId)
        dim myEmail
        set myEmail = new WaiverSearchEmail

        myEmail.SendByHashId WaiverHashId
    end sub



    '***********************************************************
    ' get the signature for the player
    '***********************************************************
    private sub WritePlayerSignature(HashId)
        dim myWaiver
        set myWaiver = new Waiver
        myWaiver.Load HashId

        Response.Write myWaiver.Signature
    end sub




%>