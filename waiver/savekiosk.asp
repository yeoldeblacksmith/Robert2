
<!--#include file="../classes/includelist.asp"-->
<%
    Err.Clear

    if len(request.Form) = 0 and len(request.QueryString) = 0 then response.Redirect "kiosk.asp"

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
                    response.Redirect("error.asp") 
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
                                response.Redirect("error.asp") 
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

    response.Redirect("kioskthanks.asp")

    function GetEventId()
        dim id
        if len(request.Form("txtEventId")) > 0 then
            id = request.Form("txtEventId")
        else
            if request.Form("ddlGrouped") = "yes" then
                id = request.Form("ddlEventGroup")
           end if
        end if
        if id = "-1" or id = "0" then id = ""
    
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
                response.Redirect("error.asp")
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
                response.Redirect("error.asp")
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
                response.Redirect("error.asp")
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
                    response.Redirect("error.asp")
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
                response.Redirect("error.asp")
            end if

            adultHash = .HashId
            
            SavePlayDateTime()
        end with

        For each itm in Request.Form
            if left(itm,6) = "Custom" AND Request.form(itm) <> "" AND (Mid(itm,InStr(itm,"_")+1,4) = "Self" OR Mid(itm,InStr(itm,"_")+1,4) = "Gene") Then
                SaveWaiverCustomField adultWaiver.WaiverId, Mid(itm,12,((InStr(itm,"_")-1)-11)), Request.form(itm)
            end if
        next
    end sub

    sub SaveMinorWaiver(MinorIndex)
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
                response.Redirect("error.asp")
            end if
    
            select case MinorIndex
               case 1:
                    minor1Hash = .HashId
                    minor1WaiverId = .WaiverId
                case 2:
                    minor2Hash = .HashId
                    minor2WaiverId = .WaiverId
               case 3:
                   minor3Hash = .HashId
                    minor3WaiverId = .WaiverId
                case 4:
                    minor4Hash = .HashId
                    minor4WaiverId = .WaiverId
                case 5:
                    minor5Hash = .HashId
                    minor5WaiverId = .WaiverId
                case 6:
                    minor6Hash = .HashId
                    minor6WaiverId = .WaiverId
            end select

            SavePlayDateTime()

        end with

        For each itm in Request.Form
            if left(itm,6) = "Custom" AND Request.form(itm) <> "" AND (Mid(itm,InStr(itm,"_")+1,6) = "Minor" & MinorIndex OR Mid(itm,InStr(itm,"_")+1,4) = "Gene") Then
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

    sub SetGuardianInfo()
        parentFirstName = request.Form("txtSelfFirstName")
        parentLastName = request.Form("txtSelfLastName")
        address = request.Form("txtSelfAddress")
        city = request.Form("txtSelfCity")
        state = request.Form("ddlSelfState")
        zipCode = request.Form("txtSelfZipCode")
        phoneNumber = request.Form("txtSelfPhoneNumber")
        emailAddress = request.Form("txtSelfEmailAddress")
    end sub
%>