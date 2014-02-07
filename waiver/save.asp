
<!--#include file="../classes/includelist.asp"-->
<%
    Err.Clear

    if len(request.Form) = 0 and len(request.QueryString) = 0 then response.Redirect "default.asp"

    dim parentFirstName, parentLastName, address, _
        city, state, zipCode, _
        eventId, playDate, playTime, _
        emailList, signature, phoneNumber, _
        emailAddress, myEvent, adultHash, _
        minor1Hash, minor2Hash, minor3Hash, _
        minor4Hash, minor5Hash, minor6Hash, _
        useRegistration, ipAddress, nPlayerId, nParentId, nWaiverHId, newAdultPlayer, searchResults, _
        msgAdult, msgMinor1, msgMinor2, msgMinor3, msgMinor4, msgMinor5, msgMinor6
    set myEvent = new ScheduledEvent

    useRegistration = cbool(request.Form("txtUseRegistration"))

    eventId = GetEventId()
    SetEventInfo()
    'emailList = iif(request.Form("chkEmailList") = "on", true, false)
    emailList = true
    signature = request.Form("txtSignature")
    ipAddress = request.ServerVariables("REMOTE_ADDR")

    If request.form("txtCheckingIn") = "true" Then
        nPlayerId = request.form("txtCheckInPID")
        nWaiverHId = GetValidWaiverHIdByPlayerId(nPlayerId)
        SavePlayDateTime()    
    Else
        If request.form("txtSearchOverride") = "true" Then
            SaveAdultPlayer()
            nWaiverHId = ""
        Else
        'check to see if adult player already has a player record and waiver status
            nPlayerId = ""
            nParentId = ""
            nWaiverHId = ""
            searchResults = GetPlayerIdWaiverStatusByNameDOBAddress(request.Form("txtSelfFirstName"), request.Form("txtSelfLastName"), cdate(request.Form("ddlSelfDOBMonth") & "/" & request.Form("ddlSelfDOBDay") & "/" & request.Form("ddlSelfDOBYear")), request.Form("txtSelfAddress"))

            Select Case Left(searchResults,1)
                Case "1"
                    'found 1 player with valid waiver
                    nPlayerId = Right(searchResults,Len(searchResults)-1)
                    nParentId = Right(searchResults,Len(searchResults)-1)
                    nWaiverHId =  GetValidWaiverHIdByPlayerId(nPlayerId)
                    msgAdult = "fail:PID" & nPlayerId & "WID" & nWaiverHId
                Case "2"
                    'found multiple players with same name, dob, address
                    On Error Resume Next
                    Err.Raise 5 
                    dim oASPErrorMulti : Set oASPErrorMulti = Server.GetLastError
                    SendExceptionMail err.number, "Multiple players found for same name, dob, address on New Waiver Save", err.Source, oAspErrorMulti
                    Err.Clear
                    On error goto 0
                    response.Redirect("error.asp?src=multi2") 
                Case "3"
                    'found 1 player with  invalid waiver    
                    nPlayerId = Right(searchResults,Len(searchResults)-1)
                    nParentId = Right(searchResults,Len(searchResults)-1)
                    nWaiverHId = ""
                Case "4"
                    'found 1 player with no previous waiver
                    nPlayerId = Right(searchResults,Len(searchResults)-1)
                    nParentId = Right(searchResults,Len(searchResults)-1)
                    nWaiverHId = ""
                Case else
                    'no player or waiver found using address
                    'search again without the address
                    searchResults = GetPlayerIdWaiverStatusByNameDOB(request.Form("txtSelfFirstName"), request.Form("txtSelfLastName"), cdate(request.Form("ddlSelfDOBMonth") & "/" & request.Form("ddlSelfDOBDay") & "/" & request.Form("ddlSelfDOBYear")))

                    Select Case Left(searchResults,1)
                        Case "1"
                            'found 1 player with valid waiver
                            nPlayerId = Right(searchResults,Len(searchResults)-1)
                            nParentId = Right(searchResults,Len(searchResults)-1)
                            nWaiverHId =  GetValidWaiverHIdByPlayerId(nPlayerId)
                            msgAdult = "fail:PID" & nPlayerId & "WID" & nWaiverHId
                        Case "2"
                            'found multiple players with same name, dob
                            'fuzzy match - flag for confirm.asp to push to search
                            nPlayerId = ""
                            nParentId = ""
                            nWaiverHId = ""
                            msgAdult = "fail:PID" & request.Form("txtSelfFirstName") & " " & request.Form("txtSelfLastName") & "WIDMissing"
                        Case "3"
                            'found 1 player with  invalid waiver    
                            nPlayerId = Right(searchResults,Len(searchResults)-1)
                            nParentId = Right(searchResults,Len(searchResults)-1)
                            nWaiverHId = ""
                        Case "4"
                            'found 1 player with no previous waiver
                            nPlayerId = Right(searchResults,Len(searchResults)-1)
                            nParentId = Right(searchResults,Len(searchResults)-1)
                            nWaiverHId = ""
                        Case else
                            'no player or waiver found 
                            'do nothing - process new waiver like before
                            SaveAdultPlayer()
    '                        response.write "saving new adult player<br />"
                            nWaiverHId = ""
                    End Select
            End Select
        End If

        if request.Form("txtSelectedPath") = 0 or request.Form("txtSelectedPath") = 2 then
            if nPlayerId <> "" AND nWaiverHId = "" Then
                SaveAdultWaiver()
                msgAdult = "pass:PID" & nPlayerId & "WID" & AdultHash
            end if
            if nPlayerId <> "" AND nWaiverHId <> "" Then
                SavePlayDateTime()
            end if
        else
            msgAdult = "pass:PID" & nPlayerId & "WIDIrrelevant"
        end if
    
        if request.Form("txtSelectedPath") = 1 or request.Form("txtSelectedPath") = 2 then     
            dim minors
            minors = cint(request.Form("txtNumberOfMinors"))
        
            if (IsEmpty(newAdultPlayer) or IsNull(newAdultPlayer)) AND (Len(nParentId) > 0)  Then
                set newAdultPlayer = new Player
                newAdultPlayer.LoadById nParentId
            end if 

            for i = 1 to minors
                    'check to see if minor player already has a player record and waiver status
                        nPlayerId = ""
                        nWaiverHId = ""
                        searchResults = GetPlayerIdWaiverStatusByNameDOBAddress(request.Form("txtMinor" & i & "FirstName"), request.Form("txtMinor" & i & "LastName"), cdate(request.Form("ddlMinor" & i & "DOBMonth") & "/" & request.Form("ddlMinor" & i & "DOBDay") & "/" & request.Form("ddlMinor" & i & "DOBYear")), request.Form("txtSelfAddress"))

                        Select Case Left(searchResults,1)
                            Case "1"
                                'found 1 player with valid waiver
                                nPlayerId = Right(searchResults,Len(searchResults)-1)
                                nWaiverHId =  GetValidWaiverHIdByPlayerId(nPlayerId)
                                Eval(msgMinor & i = "fail:PID" & nPlayerId & "WID" & nWaiverHId)
                            Case "2"
                                'found multiple players with same name, dob, address
                                On Error Resume Next
                                Err.Raise 5 
                                Set oASPErrorMulti = Server.GetLastError
                                SendExceptionMail err.number, "Multiple players found for same name, dob, address on New Waiver Save", err.Source, oAspErrorMulti
                                Err.Clear
                                On error goto 0
                                response.Redirect("error.asp?src=multi1") 
                            Case "3"
                                'found 1 player with  invalid waiver    
                                nPlayerId = Right(searchResults,Len(searchResults)-1)
                                nWaiverHId = ""
                            Case "4"
                                'found 1 player with no previous waiver
                                nPlayerId = Right(searchResults,Len(searchResults)-1)
                                nWaiverHId = ""
                            Case else
                                'no player or waiver found using address
                                'search again without the address
                                searchResults = GetPlayerIdWaiverStatusByNameDOB(request.Form("txtMinor" & i & "FirstName"), request.Form("txtMinor" & i & "LastName"), cdate(request.Form("ddlMinor" & i & "DOBMonth") & "/" & request.Form("ddlMinor" & i & "DOBDay") & "/" & request.Form("ddlMinor" & i & "DOBYear")))

                                Select Case Left(searchResults,1)
                                    Case "1"
                                        'found 1 player with valid waiver
                                        nPlayerId = Right(searchResults,Len(searchResults)-1)
                                        nWaiverHId =  GetValidWaiverHIdByPlayerId(nPlayerId)
                                        Eval(msgMinor & i = "fail:PID" & nPlayerId & "WID" & nWaiverHId)
                                    Case "2"
                                        'found multiple players with same name, dob
                                        'fuzzy match - flag for confirm.asp to push to search
                                        nPlayerId = ""
                                        nWaiverHId = ""
                                        Eval(msgMinor & i = "fail:PID" & request.Form("txtMinor" & i & "FirstName") & " " & request.Form("txtMinor" & i & "LastName") & "WIDMissing")
                                    Case "3"
                                        'found 1 player with  invalid waiver    
                                        nPlayerId = Right(searchResults,Len(searchResults)-1)
                                        nWaiverHId = ""
                                    Case "4"
                                        'found 1 player with no previous waiver
                                        nPlayerId = Right(searchResults,Len(searchResults)-1)
                                        nWaiverHId = ""
                                    Case else
                                        'no player or waiver found 
                                        'do nothing - process new waiver like before
                                        SaveMinorPlayer(i)
    '                                    response.write "saving new minor player<br />"
                                        nWaiverHId = ""
                                End Select
                        End Select

                        if nPlayerId <> "" AND nWaiverHId = "" Then
                            If Len(nParentId) > 0 Then
    '                            response.write "saving new minor waiver<br />"
                                SaveMinorWaiver(i)
                                populateMinorMsg "pass", i, nPlayerId, Eval("minor" & i & "hash") 
                            else
                                populateMinorMsg "fail", i, nPlayerId, "NoParentFound" 
                            end if
                        end if
                        if nPlayerId <> "" AND nWaiverHId <> "" Then
                            SavePlayDateTime()
                        end if
            next
        end if
    End If
%>

<html>
    <head>
        <title></title>
        <meta http-equiv="cache-control" content="max-age=0" /> 
        <meta http-equiv="cache-control" content="no-cache" /> 
        <meta http-equiv="expires" content="0" /> 
        <meta http-equiv="expires" content="Tue, 01 Jan 1980 1:00:00 GMT" /> 
        <meta http-equiv="pragma" content="no-cache" />
    </head>
    <body>
        <form action="confirm.asp" method="post">
            <input type="hidden" id="txtSelectedPath" name="txtSelectedPath" value="<%= request.Form("txtSelectedPath") %>"  />           
            <input type="hidden" id="txtAdultHash" name="txtAdultHash" value="<%= AdultHash %>" />
            <input type="hidden" id="txtMinor1Hash" name="txtMinor1Hash" value="<%= Minor1Hash %>" />
            <input type="hidden" id="txtMinor2Hash" name="txtMinor2Hash" value="<%= Minor2Hash %>" />
            <input type="hidden" id="txtMinor3Hash" name="txtMinor3Hash" value="<%= Minor3Hash %>" />
            <input type="hidden" id="txtMinor4Hash" name="txtMinor4Hash" value="<%= Minor4Hash %>" />
            <input type="hidden" id="txtMinor5Hash" name="txtMinor5Hash" value="<%= Minor5Hash %>" />
            <input type="hidden" id="txtMinor6Hash" name="txtMinor6Hash" value="<%= Minor6Hash %>" />
            <input type="hidden" id="txtMsgAdult" name="txtMsgAdult" value="<%= msgAdult %>" />
            <input type="hidden" id="txtMsgMinor1" name="txtMsgMinor1" value="<%= msgMinor1 %>" />
            <input type="hidden" id="txtMsgMinor2" name="txtMsgMinor2" value="<%= msgMinor2 %>" />
            <input type="hidden" id="txtMsgMinor3" name="txtMsgMinor3" value="<%= msgMinor3 %>" />
            <input type="hidden" id="txtMsgMinor4" name="txtMsgMinor4" value="<%= msgMinor4 %>" />
            <input type="hidden" id="txtMsgMinor5" name="txtMsgMinor5" value="<%= msgMinor5 %>" />
            <input type="hidden" id="txtMsgMinor6" name="txtMsgMinor6" value="<%= msgMinor6 %>" />
            <input type="hidden" id="txtCheckingIn" name="txtCheckingIn" value="<%= request.form("txtCheckingIn") %>" />
            <input type="hidden" id="txtCheckInPID" name="txtCheckInPID" value="<%= nPlayerId %>" />
            <input type="hidden" id="txtCheckInWID" name="txtCheckInWID" value="<%= nWaiverHId %>" />
        </form>
        <script type="text/javascript">
            document.forms[0].submit();
        </script>
    </body>
</html>
<%
    function GetEventId()
        dim id
        
        if useRegistration then
            if len(request.Form("txtEventId")) > 0 then
                id = request.Form("txtEventId")
            else
                if request.Form("ddlGrouped") = "yes" then
                    id = request.Form("ddlEventGroup")
               end if
            end if
            if id = "-1" or id = "0" then id = ""
        else
            id = ""
        end if
    
        GetEventId = id
    end function

    sub SaveWaiverCustomField(wid, wcfid, val)
        dim myCustomWaiverField
        set myCustomWaiverField = new WaiverCustomValue

        with myCustomWaiverField
            .WaiverId = wid
            .WaiverCustomFieldId = wcfid
            .Value = val

            .Save

            if Err.Number <> 0 then
                dim oASPError1 : Set oASPError1 = Server.GetLastError
                    SendExceptionMail err.number, err.Description, err.Source, oAspError1
                    response.Redirect("error.asp?src=swcf1")
            end if
        end with

    end sub

    function GetValidWaiverHIdByPlayerId(playerid)
        dim myWaivers, foundWaiver, returnValue
        returnValue = ""
        set myWaivers = new WaiverCollection
        foundWaiver = false

        myWaivers.LoadWaiversByPlayerId playerId

        for i = 0 to myWaivers.Count -1
            if myWaivers(i).ValidityStatus = WAIVER_VALIDITYSTATUS_GOOD then
                returnValue = myWaivers(i).HashId
                foundWaiver = true
            end if
        next

        if NOT foundWaiver then
            returnValue = "waiver not found"
        end if

        GetValidWaiverHIdByPlayerId = returnValue
    end function

    function GetPlayerIdWaiverStatusByNameDOB(sFirstName, sLastName, dtDOB)
        dim players, waivers, foundGoodWaiver, foundInvalidWaiver, returnValue
        returnValue = ""
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
                    returnValue = "1" & players(0).playerId
                else
                    if foundInvalidWaiver then
                        returnValue = "3" & players(0).playerId
                    else
                        returnValue = "4" & players(0).playerId
                    end if
                end if
            Case players.Count >= 2
                    returnValue = "2"
                    for x = 0 to players.count -1
                        returnValue = returnValue & "@" & players(x).playerId
                    next
            Case Else
                    returnValue = "no players found"
        End Select

        GetPlayerIdWaiverStatusByNameDOB = returnValue
    end function


    function GetPlayerIdWaiverStatusByNameDOBAddress(sFirstName, sLastName, dtDOB, sAddress)
        dim players, waivers, foundGoodWaiver, foundInvalidWaiver, returnValue
        returnValue = ""
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
                        returnValue = "1" & players(0).playerId
                    else
                        if foundInvalidWaiver then
                            returnValue = "3" & players(0).playerId
                        else
                            returnValue = "4" & players(0).playerId
                        end if
                    end if
   
            Case players.Count >= 2
                    returnValue = "2"
                    for x = 0 to players.count -1
                        returnValue = returnValue & "@" & players(x).playerId
                    next
            Case Else
                    returnValue = "no players found"
        End Select

        GetPlayerIdWaiverStatusByNameDOBAddress = returnValue
    end function

    sub populateMinorMsg(passfail, minorcnt, currentPlayerId, currentWaiverId)
        Select Case minorcnt
            case 1
                msgMinor1 = passfail & ":PID" & currentPlayerId & "WID" & currentWaiverId
            case 2
                msgMinor2 = passfail & ":PID" & currentPlayerId & "WID" & currentWaiverId
            case 3
                msgMinor3 = passfail & ":PID" & currentPlayerId & "WID" & currentWaiverId
            case 4
                msgMinor4 = passfail & ":PID" & currentPlayerId & "WID" & currentWaiverId
            case 5
                msgMinor5 = passfail & ":PID" & currentPlayerId & "WID" & currentWaiverId
            case 6
                msgMinor6 = passfail & ":PID" & currentPlayerId & "WID" & currentWaiverId
        End Select
    end sub

    sub SaveAdultPlayer()
        'on error resume next
        set newAdultPlayer = new Player

        with newAdultPlayer
            .PlayerId = 0
            .EmailList = emailList
            .DateOfBirth = cdate(request.Form("ddlSelfDOBMonth") & "/" & request.Form("ddlSelfDOBDay") & "/" & request.Form("ddlSelfDOBYear"))          
            .City = request.Form("txtSelfCity")
            .State = request.Form("ddlSelfState")
            .ZipCode = request.Form("txtSelfZipCode")
            .PhoneNumber = request.Form("txtSelfPhoneNumber")
            .FirstName = request.Form("txtSelfFirstName")
            .LastName = request.Form("txtSelfLastName")
            .Address = request.Form("txtSelfAddress")
            .EmailAddress = request.Form("txtSelfEmailAddress")

            .Save

            if Err.Number <> 0 then
                dim oASPError : Set oASPError = Server.GetLastError
                SendExceptionMail err.number, err.Description, err.Source, oAspError
                response.Redirect("error.asp?src=snadplyr")
            end if

            nPlayerId = .PlayerId
            nParentId = .PlayerId

        end with

    end sub

    sub SaveMinorPlayer(minorcnt)
        'on error resume next

        dim newMinorPlayer
        set newMinorPlayer = new Player
        
        with newMinorPlayer
            .PlayerId = 0
            .EmailList = emailList
            .DateOfBirth = cdate(request.Form("ddlMinor" & minorcnt & "DOBMonth") & "/" & request.Form("ddlMinor" & minorcnt & "DOBDay") & "/" & request.Form("ddlMinor" & minorcnt & "DOBYear"))          
            .City = newAdultPlayer.City
            .State = newAdultPlayer.State
            .ZipCode = newAdultPlayer.ZipCode
            .PhoneNumber = newAdultPlayer.PhoneNumber
            .FirstName = request.Form("txtMinor" & minorcnt & "FirstName")
            .LastName = request.Form("txtMinor" & minorcnt & "LastName")
            .Address = newAdultPlayer.Address
            .EmailAddress = newAdultPlayer.EmailAddress

            .Save

            if Err.Number <> 0 then
                dim oASPError : Set oASPError = Server.GetLastError
                SendExceptionMail err.number, err.Description, err.Source, oAspError
                response.Redirect("error.asp?src=snmplr")
            end if

            nPlayerId = .PlayerId

        end with

    end sub

    sub SavePlayDateTime()

        'on error resume next

        dim newPlayDateTime, myPlayDates, originalPlayTime
        set newPlayDateTime = new PlayDateTime
        set myPlayDates = new PlayDateTimeCollection
 
        myPlayDates.LoadAllByPlayerAndDate nPlayerId, playDate

        if myPlayDates.Count > 0 Then
            for ndx = 0 to myPlayDates.Count -1
                originalPlayTime = myPlayDates(ndx).PlayTime
                myPlayDates(ndx).PlayTime = playTime
                myPlayDates(ndx).EventId = EventId
                myPlayDates(ndx).UpdatePlayDateEventAndTime originalPlayTime
           next            
        else
            with newPlayDateTime
                .PlayerId = nPlayerId
                .PlayDate = playDate
                .PlayTime = playTime        
                .EventId = EventId

                .Save

                if Err.Number <> 0 then
                    dim oASPError : Set oASPError = Server.GetLastError
                    SendExceptionMail err.number, err.Description, err.Source, oAspError
                    response.Redirect("error.asp?src=spdt")
                end if
            end with
        end If
    end sub

    sub SaveAdultWaiver()
        'on error resume next
        dim adultWaiver
        set adultWaiver = new Waiver

        with adultWaiver
            .WaiverId = 0
            .PlayerId = nPlayerId
            .PhoneNumber = request.Form("txtSelfPhoneNumber")
            .Address = request.Form("txtSelfAddress")
            .City = request.Form("txtSelfCity")
            .State = request.Form("ddlSelfState")
            .ZipCode = request.Form("txtSelfZipCode")
            .CreateDateTime = Now()
            .EmailAddress = request.Form("txtSelfEmailAddress")
            .SubmittedFromIpAddress = ipAddress 
            .Signature = Signature

            .Save        

            if Err.Number <> 0 then
                dim oASPError1 : Set oASPError1 = Server.GetLastError
                SendExceptionMail err.number, err.Description, err.Source, oAspError1
                response.Redirect("error.asp?src=sadwvr")
            end if

            adultHash = .HashId
            
            SavePlayDateTime()
        end with

        For each itm in Request.Form
            if left(itm,6) = "Custom" AND (Mid(itm,InStr(itm,"_")+1,4) = "Self" OR Mid(itm,InStr(itm,"_")+1,4) = "Gene") Then
                SaveWaiverCustomField adultWaiver.WaiverId, Mid(itm,12,((InStr(itm,"_")-1)-11)), Request.form(itm)
            end if
        next
    end sub

    sub SaveMinorWaiver(MinorIndex)
        on error resume next
        dim minorWaiver
        set minorWaiver = new Waiver

        with minorWaiver
            .WaiverId = 0
            .PlayerId = nPlayerId
            .ParentId = nParentId
            .PhoneNumber = newAdultPlayer.phoneNumber
            .Address = newAdultPlayer.address
            .City = newAdultPlayer.city
            .State = newAdultPlayer.state
            .ZipCode = newAdultPlayer.zipCode
            .CreateDateTime = Now()
            .EmailAddress = newAdultPlayer.emailAddress
            .SubmittedFromIpAddress = ipAddress 
            .Signature = Signature
    
            .Save

            if Err.Number <> 0 then
                dim oASPError1 : Set oASPError1 = Server.GetLastError
                SendExceptionMail err.number, err.Description, err.Source, oAspError1
                response.Redirect("error.asp?src=smwvr")
            end if

            select case MinorIndex
               case 1:
                    minor1Hash = .HashId
                case 2:
                    minor2Hash = .HashId
               case 3:
                   minor3Hash = .HashId
                case 4:
                    minor4Hash = .HashId
                case 5:
                    minor5Hash = .HashId
                case 6:
                    minor6Hash = .HashId
            end select

            SavePlayDateTime()

        end with

        For each itm in Request.Form
            if left(itm,6) = "Custom" AND (Mid(itm,InStr(itm,"_")+1,6) = "Minor" & MinorIndex OR Mid(itm,InStr(itm,"_")+1,4) = "Gene") Then
                SaveWaiverCustomField minorWaiver.WaiverId, Mid(itm,12,((InStr(itm,"_")-1)-11)), Request.form(itm)
            end if
        next
    end sub

    sub SendExceptionMail(ErrNumber, ErrDescription, ErrSource, AspError)
        dim exEmail : set exEmail = new ExceptionEmail
        exEmail.Send Request, ErrNumber, ErrDescription, ErrSource, AspError
    end sub

    sub SetEventInfo()
        if len(eventId) = 0 then
            playDate = request.Form("txtPlayDate")
            playTime = request.Form("ddlPlayTime")
        else
            myEvent.Load eventId
            playDate = myEvent.SelectedDate
            playTime = myEvent.StartTimeShortFormat
        end if
    end sub

%>